VERSION 5.00
Begin VB.Form dialog_GradientEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gradient editor"
   ClientHeight    =   7995
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
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   844
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   5895
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   10398
      Begin PhotoDemon.pdCheckBox chkDistributeEvenly 
         Height          =   330
         Left            =   360
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   2400
         Width           =   12615
      End
      Begin VB.PictureBox picNodePreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1950
         Left            =   240
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   807
         TabIndex        =   7
         Top             =   360
         Width           =   12135
      End
      Begin PhotoDemon.pdSlider sltNodeOpacity 
         Height          =   705
         Left            =   4320
         TabIndex        =   5
         Top             =   3660
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   1270
         Caption         =   "opacity"
         Max             =   100
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.pdColorSelector csNode 
         Height          =   855
         Left            =   240
         TabIndex        =   4
         Top             =   3660
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   1508
         Caption         =   "color"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   0
         Left            =   120
         Top             =   3240
         Width           =   12135
         _ExtentX        =   16536
         _ExtentY        =   556
         Caption         =   "current node settings"
         FontSize        =   12
      End
      Begin PhotoDemon.pdSlider sltNodePosition 
         Height          =   705
         Left            =   8280
         TabIndex        =   6
         Top             =   3660
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
         Index           =   2
         Left            =   120
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
         Top             =   2880
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
         Index           =   4
         Left            =   120
         Top             =   4800
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   556
         Caption         =   "additional tools"
         FontSize        =   12
      End
   End
   Begin PhotoDemon.pdButtonStrip btsEdit 
      Height          =   915
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   1614
      Caption         =   "edit mode"
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   7245
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   1323
      AutoloadLastPreset=   -1  'True
      DontAutoUnloadParent=   -1  'True
      DontResetAutomatically=   -1  'True
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4095
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7223
   End
End
Attribute VB_Name = "dialog_GradientEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gradient Editor Dialog
'Copyright 2014-2016 by Tanner Helland
'Created: 23/July/15 (but assembled from many bits written earlier)
'Last updated: 23/July/15
'Last update: initial build
'
'Comprehensive gradient editor.  This dialog is currently based around the properties of GDI+ gradient brushes, but it
' could easily be expanded in the future due to its modular design.
'
'Note that - by design - this editor always returns a gradient with the same shape and angle as it was passed.  This editor
' does not allow you to set the gradient's shape and/or angle.  (That must be done externally.)
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

'Gradient strings are generated with the help of PD's core gradient class.  Please note that within this dialog, the gradient's
' shape is ignored; only a linear gradient is displayed, to make editing easier.
Private m_NodePreview As pd2DGradient

'If a user control spawned this dialog, it will pass itself as a reference.  We can then send gradient updates back
' to the control, allowing for real-time updates on the screen despite a modal dialog being raised!
Private parentGradientControl As pdGradientSelector

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
Private m_GradientPoints() As GRADIENTPOINT

'The current gradient point (index) selected and/or hovered by the mouse.  -1 if no point is currently selected/hovered.
Private m_CurPoint As Long, m_CurHoverPoint As Long

'Size of gradient "nodes" in the interactive UI.
Private Const GRADIENT_NODE_WIDTH As Single = 12#
Private Const GRADIENT_NODE_HEIGHT As Single = 14#

'Other gradient node UI renderers
Private inactiveArrowFill As pd2DBrush, activeArrowFill As pd2DBrush
Private inactiveOutlinePen As pd2DPen, activeOutlinePen As pd2DPen

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The newly selected gradient (if any) is returned via this property
Public Property Get NewGradient() As String
    NewGradient = GetGradientAsOriginalShape
End Property

'This dialog is a little confusing because it *only* operates on linear gradients.  If it's passed something like a radial gradient,
' it will combine the original shape and/or angle it was passed with the current node settings to arrive at a new gradient string.
Private Function GetGradientAsOriginalShape() As String
    
    Dim initGradient As pd2DGradient
    Set initGradient = New pd2DGradient
    initGradient.CreateGradientFromString m_OldGradient
    
    Dim tmpGradient As pd2DGradient
    Set tmpGradient = New pd2DGradient
    tmpGradient.CreateGradientFromString m_NodePreview.GetGradientAsString
    tmpGradient.SetGradientShape initGradient.GetGradientShape
    tmpGradient.SetGradientAngle initGradient.GetGradientAngle
    tmpGradient.SetGradientWrapMode initGradient.GetGradientWrapMode
    
    GetGradientAsOriginalShape = tmpGradient.GetGradientAsString
    
End Function

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByVal initialGradient As String, Optional ByRef callingControl As pdGradientSelector = Nothing)
    
    'Store a reference to the calling control (if any)
    Set parentGradientControl = callingControl

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Cache the initial gradient parameters so we can access it elsewhere
    m_OldGradient = initialGradient
    
    'Inside this dialog, the gradient is always forced to a linear-type gradient at angle 0.  This makes it much easier to edit.
    Set m_NodePreview = New pd2DGradient
    m_NodePreview.CreateGradientFromString m_OldGradient
    m_NodePreview.SetGradientShape P2_GS_Linear
    m_NodePreview.SetGradientAngle 0#
    
    'If the dialog is being initialized for the first time, there will be no "initial gradient".  In this case, the gradient class
    ' will initialize a placeholder gradient.  We make a copy of it, and use that as the basis of the editor's initial settings.
    If Len(m_OldGradient) = 0 Then m_OldGradient = m_NodePreview.GetGradientAsString
    
    'Sync all controls to the initial pen parameters
    SyncControlsToGradientObject
    UpdatePreview
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    
    'Apply extra images and tooltips to certain controls
    
    'Apply visual themes
    ApplyThemeAndTranslations Me
    
    'Initialize an XML engine, which we will use to read/write recent pen data to file
    Set m_XMLEngine = New pdXML
    
    'The XML file will be stored in the Preset path (/Data/Presets)
    m_XMLFilename = g_UserPreferences.GetPresetPath & "Gradient_Selector.xml"
    
    'TODO: if an XML file exists, load its contents now
    'LoadRecentGradientList
        
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub btsEdit_Click(ByVal buttonIndex As Long)
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        picContainer(i).Visible = CBool(i = buttonIndex)
    Next i
End Sub

Private Sub chkDistributeEvenly_Click()
    If (Not m_SuspendUI) Then
        RedrawEverything
        SyncUIToActiveNode
    End If
End Sub

Private Sub cmdBar_AddCustomPresetData()

    'This control (obviously) requires a lot of extra custom preset data.
    '
    'However, there's no reason to require horrible duplication code, when the gradient class is already capable of serializing
    ' all relevant data for this control!
    cmdBar.AddPresetData "FullGradientDefinition", GetGradientAsOriginalShape()

End Sub

'CANCEL BUTTON
Private Sub cmdBar_CancelClick()
    userAnswer = vbCancel
    Me.Hide
End Sub

'OK BUTTON
Private Sub cmdBar_OKClick()

    'Store the newGradient value (which the dialog handler will use to return the selected gradient to the caller)
    UpdateGradientObjects
    
    'TODO: save the current list of recently used gradients
    'SaveRecentGradientList
    
    userAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_ReadCustomPresetData()
    
    'This control (obviously) requires a lot of extra custom preset data.
    '
    'However, there's no reason to require horrible duplication code, when the gradient class is already capable of serializing
    ' all relevant data for this control!
    m_NodePreview.CreateGradientFromString cmdBar.RetrievePresetData("FullGradientDefinition")
    
    'Synchronize all controls to the updated settings
    SyncControlsToGradientObject
        
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Reset our master gradient object; everything else derives from it
    Set m_NodePreview = New pd2DGradient
    m_NodePreview.CreateGradientFromString ""
    
    'Synchronize all controls to the updated settings
    UpdateGradientObjects
    SyncControlsToGradientObject
    UpdatePreview
    
End Sub

Private Sub csNode_ColorChanged()
    
    If (Not m_SuspendUI) And (m_CurPoint >= 0) Then
        m_GradientPoints(m_CurPoint).PointRGB = csNode.Color
        RedrawEverything
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
    
    If g_IsProgramRunning Then
        
        If (m_NodePreview Is Nothing) Then Set m_NodePreview = New pd2DGradient
        
        'Set up a special mouse handler for the gradient interaction window
        If (m_MouseEvents Is Nothing) Then Set m_MouseEvents = New pdInputMouse
        m_MouseEvents.AddInputTracker picInteract.hWnd, True, True, , True
        m_MouseEvents.SetSystemCursor IDC_HAND
        
        'Prep a default set of gradient points
        ResetGradientPoints
        
        'Prep all gradient point tracking variables
        m_CurPoint = -1
        
        'While we're here, we'll also prep all generic drawing objects for the interactive gradient node UI bits
        Set inactiveArrowFill = New pd2DBrush
        Set activeArrowFill = New pd2DBrush
        
        inactiveArrowFill.SetBrushProperty P2_BrushMode, 0
        inactiveArrowFill.SetBrushProperty P2_BrushOpacity, 100
        inactiveArrowFill.SetBrushProperty P2_BrushColor, g_Themer.GetGenericUIColor(UI_Background)
        inactiveArrowFill.CreateBrush
        
        activeArrowFill.SetBrushProperty P2_BrushMode, 0
        activeArrowFill.SetBrushProperty P2_BrushOpacity, 100
        activeArrowFill.SetBrushProperty P2_BrushColor, g_Themer.GetGenericUIColor(UI_AccentLight)
        activeArrowFill.CreateBrush
        
        Set inactiveOutlinePen = New pd2DPen
        Set activeOutlinePen = New pd2DPen
        
        inactiveOutlinePen.SetPenProperty P2_PenStyle, GP_DS_Solid
        inactiveOutlinePen.SetPenProperty P2_PenOpacity, 100
        inactiveOutlinePen.SetPenProperty P2_PenWidth, 1#
        inactiveOutlinePen.SetPenProperty P2_PenLineJoin, GP_LJ_Miter
        inactiveOutlinePen.SetPenProperty P2_PenColor, g_Themer.GetGenericUIColor(UI_GrayDark)
        inactiveOutlinePen.CreatePen
        
        activeOutlinePen.SetPenProperty P2_PenStyle, GP_DS_Solid
        activeOutlinePen.SetPenProperty P2_PenOpacity, 100
        activeOutlinePen.SetPenProperty P2_PenWidth, 1#
        activeOutlinePen.SetPenProperty P2_PenLineJoin, GP_LJ_Miter
        activeOutlinePen.SetPenProperty P2_PenColor, g_Themer.GetGenericUIColor(UI_Accent)
        activeOutlinePen.CreatePen
                
        'Draw the initial set of interactive gradient nodes
        SyncUIToActiveNode
        DrawGradientNodes
                
    End If
    
    m_SuspendUI = False
    
End Sub

Private Sub ResetGradientPoints()
    
    m_NumOfGradientPoints = 2
    ReDim m_GradientPoints(0 To m_NumOfGradientPoints - 1) As GRADIENTPOINT
    
    With m_GradientPoints(0)
        .PointRGB = vbBlack
        .PointOpacity = 100
        .PointPosition = 0
    End With
    
    With m_GradientPoints(1)
        .PointRGB = vbWhite
        .PointOpacity = 100
        .PointPosition = 1
    End With
    
    If (m_NodePreview Is Nothing) Then Set m_NodePreview = New pd2DGradient
    m_NodePreview.CreateGradientFromPointCollection m_NumOfGradientPoints, m_GradientPoints
    
End Sub

Private Sub Form_Resize()
    DrawGradientNodes
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Update our two internal gradient classes against any/all changed settings.
' (Note that the node-editor class only reflects the current collection of colors and positions, not things like angle or gradient type,
'  so we only sync it against the node collection.)
Private Sub UpdateGradientObjects()
    
    'If the "evenly distribute nodes" option is checked, assign positions automatically.
    If CBool(chkDistributeEvenly.Value) Then
        
        'Start by sorting nodes from least-to-greatest.  This has the unintended side-effect of changing the active node, unfortunately,
        ' so we must also reset the active node (if any).
        
        'Start by seeing if nodes require sorting.
        Dim i As Long
        
        Dim sortNeeded As Boolean
        sortNeeded = False
        
        For i = 1 To m_NumOfGradientPoints - 1
            If m_GradientPoints(i).PointPosition < m_GradientPoints(i - 1).PointPosition Then
                sortNeeded = True
                Exit For
            End If
        Next i
        
        'If a sort is required, perform it now
        If sortNeeded Then
            
            m_CurPoint = -1
            m_CurHoverPoint = -1
            
            SyncUIToActiveNode
            
            m_NodePreview.CreateGradientFromPointCollection m_NumOfGradientPoints, m_GradientPoints
            m_NodePreview.GetCopyOfPointCollection m_NumOfGradientPoints, m_GradientPoints
            
        End If
        
        'Redistribute points accordingly
        For i = 0 To m_NumOfGradientPoints - 1
            m_GradientPoints(i).PointPosition = i / (m_NumOfGradientPoints - 1)
        Next i
        
    End If
    
    With m_NodePreview
        .SetGradientProperty P2_GradientShape, P2_GS_Linear
        .SetGradientProperty P2_GradientAngle, 0#
        .CreateGradientFromPointCollection m_NumOfGradientPoints, m_GradientPoints
    End With

End Sub

Private Sub UpdatePreview()
    
    If (Not m_SuspendUI) Then
    
        'Make sure our gradient objects are up-to-date
        UpdateGradientObjects
        
        'Next, use the current gradient nodes to paint a matching preview across the node editor window
        Dim boundsRect As RECTF, gdipBrushNodes As Long
        
        With boundsRect
            .Left = 0
            .Top = 0
            .Width = picNodePreview.ScaleWidth
            .Height = picNodePreview.ScaleHeight
        End With
        
        gdipBrushNodes = m_NodePreview.GetBrushHandle(boundsRect, True)
        
        If (m_NodePreviewDIB Is Nothing) Then Set m_NodePreviewDIB = New pdDIB
        If (m_NodePreviewDIB.GetDIBWidth <> Me.picNodePreview.ScaleWidth) Or (m_NodePreviewDIB.GetDIBHeight <> Me.picNodePreview.ScaleHeight) Then
            m_NodePreviewDIB.CreateBlank Me.picNodePreview.ScaleWidth, Me.picNodePreview.ScaleHeight, 24, 0
        Else
            m_NodePreviewDIB.ResetDIB
        End If
        
        With m_NodePreviewDIB
            GDI_Plus.GDIPlusFillDIBRect_Pattern m_NodePreviewDIB, 0, 0, .GetDIBWidth, .GetDIBHeight, g_CheckerboardPattern
            GDI_Plus.GDIPlusFillDC_Brush .GetDIBDC, gdipBrushNodes, 0, 0, .GetDIBWidth, .GetDIBHeight
        End With
        
        m_NodePreviewDIB.RenderToPictureBox Me.picNodePreview
        
        'Release our GDI+ handles
        GDI_Plus.ReleaseGDIPlusBrush gdipBrushNodes
                
        'Notify our parent of the update
        If Not (parentGradientControl Is Nothing) Then parentGradientControl.NotifyOfLiveGradientChange GetGradientAsOriginalShape()
        
    End If
    
End Sub

'Make all UI elements reflect the current gradient object.  This is typically done after the dialog loads, or after loading a
' previously created gradient.
Private Sub SyncControlsToGradientObject()
        
    m_SuspendUI = True
    
    With m_NodePreview
        .GetCopyOfPointCollection m_NumOfGradientPoints, m_GradientPoints
    End With
    
    DrawGradientNodes
    
    m_SuspendUI = False
    
    'Also, synchronize the node-specific UI to the active node (if any)
    SyncUIToActiveNode
    
End Sub

Private Sub m_MouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    Dim i As Long
    
    'Clicking the mouse either selects an existing point, or creates a new point.
    ' As such, this function will always result in a legitimate value for m_CurPoint.
    
    'See if an existing has been selected.
    Dim tmpPoint As Long
    tmpPoint = GetPointAtPosition(x, y)
    
    'If this is an existing point, we will either (LMB) mark it as the active point, or (RMB) remove it
    If (tmpPoint >= 0) Then
        
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
                SyncUIToActiveNode
            End If
            
        End If
        
    'If this is not an existing point, create a new one now.
    Else
    
        'Enlarge the target array as necessary
        If m_NumOfGradientPoints >= UBound(m_GradientPoints) Then ReDim Preserve m_GradientPoints(0 To m_NumOfGradientPoints * 2) As GRADIENTPOINT
        
        With m_GradientPoints(m_NumOfGradientPoints)
            .PointOpacity = 100
            .PointPosition = ConvertPixelCoordsToNodeCoords(x)
            
            'Preset the RGB value to match whatever the gradient already is at this point
            Dim newRGBA As RGBQUAD
            m_NodePreview.GetColorAtPosition_RGBA .PointPosition, newRGBA
            .PointRGB = RGB(newRGBA.Red, newRGBA.Green, newRGBA.Blue)
            
        End With
        
        m_CurPoint = m_NumOfGradientPoints
        m_NumOfGradientPoints = m_NumOfGradientPoints + 1
        
    End If
    
    'Regardless of outcome, we need to resync the UI to the active node, and redraw the interaction area and preview
    SyncUIToActiveNode
    UpdatePreview
    DrawGradientNodes

End Sub

Private Sub m_MouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseEvents.SetSystemCursor IDC_HAND
End Sub

Private Sub m_MouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_CurHoverPoint = -1
    m_MouseEvents.SetSystemCursor IDC_DEFAULT
    DrawGradientNodes
End Sub

Private Sub m_MouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'First, separate our handling by mouse button state
    If (Button And pdLeftButton) <> 0 Then
    
        'The left mouse button is down.  Assign the new position to the active node.
        If (m_CurPoint >= 0) Then
            If CBool(chkDistributeEvenly.Value) Then chkDistributeEvenly.Value = vbUnchecked
            m_GradientPoints(m_CurPoint).PointPosition = ConvertPixelCoordsToNodeCoords(x)
        End If
        
        'Redraw the gradient interaction nodes and the gradient itself
        SyncUIToActiveNode
        DrawGradientNodes
        UpdatePreview
        
    'The left mouse button is not down
    Else
    
        'See if a new point is currently being hovered.
        Dim tmpPoint As Long
        tmpPoint = GetPointAtPosition(x, y)
        
        'If a new point is being hovered, highlight it and redraw the interactive area
        If tmpPoint <> m_CurHoverPoint Then
            m_CurHoverPoint = tmpPoint
            DrawGradientNodes
        End If
    
    End If
    
End Sub

'Given an x-position in the interaction box, return the currently hovered point.  If multiple points are hovered, the nearest one will be returned.
Private Function GetPointAtPosition(ByVal x As Long, y As Long) As Long
    
    'Start by converting the current x-position into the range [0, 1]
    Dim convPoint As Single
    convPoint = ConvertPixelCoordsToNodeCoords(x)
    
    'convPoint now contains the position of the mouse on the range [0, 1].  Find the nearest point.
    Dim minDistance As Single, curDistance As Single, minIndex As Long
    minDistance = 1
    minIndex = -1
    
    Dim i As Long
    For i = 0 To m_NumOfGradientPoints - 1
        curDistance = Abs(m_GradientPoints(i).PointPosition - convPoint)
        If curDistance < minDistance Then
            minIndex = i
            minDistance = curDistance
        End If
    Next i
    
    'The nearest point (if any) will now be in minIndex.  If it falls below the valid threshold for clicks, accept it.
    If minDistance < (GRADIENT_NODE_WIDTH / 2) / CDbl(picNodePreview.ScaleWidth) Then
        GetPointAtPosition = minIndex
    Else
        GetPointAtPosition = -1
    End If
    
End Function

'Given an (x, y) position on the gradient interaction window, convert it to the [0, 1] range used by the gradient control.
Private Function ConvertPixelCoordsToNodeCoords(ByVal x As Long) As Single
    
    'Start by converting the current x-position into the range [0, 1]
    Dim uiMin As Single, uiMax As Single, uiRange As Single
    uiMin = scaleX(picNodePreview.Left, vbTwips, vbPixels) + 1
    uiMax = scaleX(picNodePreview.Left, vbTwips, vbPixels) + picNodePreview.ScaleWidth
    uiRange = uiMax - uiMin
    
    ConvertPixelCoordsToNodeCoords = (CSng(x) - uiMin) / uiRange
    
    If ConvertPixelCoordsToNodeCoords < 0 Then
        ConvertPixelCoordsToNodeCoords = 0
    ElseIf ConvertPixelCoordsToNodeCoords > 1 Then
        ConvertPixelCoordsToNodeCoords = 1
    End If
    
End Function

'When a new active node is selected (or its parameters somehow changed), call this sub to synchronize all UI elements to that node's properties.
Private Sub SyncUIToActiveNode()
    
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
                csNode.Color = .PointRGB
                sltNodeOpacity.Value = .PointOpacity
                sltNodePosition.Value = .PointPosition * 100
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
Private Sub DrawGradientNodes()

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
        baseArrow.AddTriangle -1 * triangleHalfWidth, triangleHeight, 0, 0, triangleHalfWidth, triangleHeight
        
        'Next up is the colored block, also centered horizontally around position 0
        baseBlock.AddRectangle_Relative -1 * GRADIENT_NODE_WIDTH \ 2, triangleHeight, GRADIENT_NODE_WIDTH, GRADIENT_NODE_HEIGHT
        
        'We also want some duplicate nodes, to remove the need to reset our base node shapes between draws
        Dim tmpArrow As pdGraphicsPath, tmpBlock As pdGraphicsPath
        Set tmpArrow = New pdGraphicsPath
        Set tmpBlock = New pdGraphicsPath
        
        'Finally, some generic scale factors to simplify the process of positioning nodes (who store their positions on the range [0, 1])
        Dim hOffset As Single, hScaleFactor As Single
        hOffset = scaleX((picNodePreview.Left - picInteract.Left), vbTwips, vbPixels) + 1
        hScaleFactor = (picNodePreview.ScaleWidth - 1)
        
        '...and pen/fill objects for the actual rendering
        Dim blockFill As pd2DBrush
        Set blockFill = New pd2DBrush
        blockFill.SetBrushProperty P2_BrushMode, 0
        blockFill.SetBrushProperty P2_BrushOpacity, 100
        
        'Prep the target interaction DIB
        If (m_InteractiveDIB Is Nothing) Then Set m_InteractiveDIB = New pdDIB
        If (m_InteractiveDIB.GetDIBWidth <> Me.picInteract.ScaleWidth) Or (m_InteractiveDIB.GetDIBHeight <> Me.picInteract.ScaleHeight) Then
            m_InteractiveDIB.CreateBlank Me.picInteract.ScaleWidth, Me.picInteract.ScaleHeight, 24, 0
        Else
            m_InteractiveDIB.ResetDIB
        End If
        
        'Fill the interaction DIB with white
        GDI_Plus.GDIPlusFillDIBRect m_InteractiveDIB, 0, 0, m_InteractiveDIB.GetDIBWidth, m_InteractiveDIB.GetDIBHeight, vbWhite, 255
        
        'Now all we do is use those to draw all the nodes in turn
        Dim i As Long
        For i = 0 To m_NumOfGradientPoints - 1
            
            'Copy the base shapes
            tmpArrow.CloneExistingPath baseArrow
            tmpBlock.CloneExistingPath baseBlock
            
            'Translate them to this node's position
            tmpArrow.TranslatePath hOffset + m_GradientPoints(i).PointPosition * hScaleFactor, 0
            tmpBlock.TranslatePath hOffset + m_GradientPoints(i).PointPosition * hScaleFactor, 0
            
            'The node's colored block is rendered the same regardless of hover
            blockFill.SetBrushProperty P2_BrushColor, m_GradientPoints(i).PointRGB
            tmpBlock.FillPathToDIB_BareBrush blockFill.GetHandle, m_InteractiveDIB
            
            'All other renders vary by hover state
            If (i = m_CurPoint) Then
                tmpBlock.StrokePath_BarePen activeOutlinePen.GetHandle, m_InteractiveDIB.GetDIBDC
                tmpArrow.FillPathToDIB_BareBrush activeArrowFill.GetHandle, m_InteractiveDIB
                tmpArrow.StrokePath_BarePen activeOutlinePen.GetHandle, m_InteractiveDIB.GetDIBDC
            ElseIf (i = m_CurHoverPoint) Then
                tmpBlock.StrokePath_BarePen activeOutlinePen.GetHandle, m_InteractiveDIB.GetDIBDC
                tmpArrow.FillPathToDIB_BareBrush activeArrowFill.GetHandle, m_InteractiveDIB
                tmpArrow.StrokePath_BarePen activeOutlinePen.GetHandle, m_InteractiveDIB.GetDIBDC
            Else
                tmpBlock.StrokePath_BarePen inactiveOutlinePen.GetHandle, m_InteractiveDIB.GetDIBDC
                tmpArrow.FillPathToDIB_BareBrush inactiveArrowFill.GetHandle, m_InteractiveDIB
                tmpArrow.StrokePath_BarePen inactiveOutlinePen.GetHandle, m_InteractiveDIB.GetDIBDC
            End If
            
        Next i
        
        'Finally, flip the DIB to the screen
        m_InteractiveDIB.RenderToPictureBox picInteract
        
    End If

End Sub

'Some user interactions require us to redraw just about everything on the dialog.  Use this shortcut function to do so.
Private Sub RedrawEverything()
    UpdateGradientObjects
    DrawGradientNodes
    UpdatePreview
End Sub

Private Sub sltNodeOpacity_Change()
    If (Not m_SuspendUI) And (m_CurPoint >= 0) Then
        m_GradientPoints(m_CurPoint).PointOpacity = sltNodeOpacity.Value
        RedrawEverything
    End If
End Sub

Private Sub sltNodePosition_Change()
    If (Not m_SuspendUI) And (m_CurPoint >= 0) Then
        m_GradientPoints(m_CurPoint).PointPosition = sltNodePosition.Value / 100
        RedrawEverything
    End If
End Sub

