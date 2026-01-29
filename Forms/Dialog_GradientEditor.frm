VERSION 5.00
Begin VB.Form dialog_GradientEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gradient editor"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12660
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   844
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdContainer pnlShared 
      Height          =   2175
      Left            =   120
      Top             =   5880
      Visible         =   0   'False
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   3836
      Begin PhotoDemon.pdTextBox txtName 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   465
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdCheckBox chkDistributeEvenly 
         Height          =   330
         Left            =   8280
         TabIndex        =   8
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         Caption         =   "make node distances equal"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   4
         Left            =   4440
         Top             =   0
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   556
         Caption         =   "additional options"
         FontSize        =   12
      End
      Begin PhotoDemon.pdCheckBox chkGamma 
         Height          =   330
         Left            =   4680
         TabIndex        =   2
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         Caption         =   "use gamma when blending"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         Caption         =   "gradient name"
         FontSize        =   12
      End
      Begin PhotoDemon.pdButton cmdFile 
         Height          =   615
         Index           =   1
         Left            =   5340
         TabIndex        =   3
         Top             =   1440
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   1085
         Caption         =   "import gradient file"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   3
         Left            =   0
         Top             =   1080
         Width           =   12135
         _ExtentX        =   16536
         _ExtentY        =   556
         Caption         =   "import / export"
         FontSize        =   12
      End
      Begin PhotoDemon.pdButton cmdFile 
         Height          =   615
         Index           =   2
         Left            =   8880
         TabIndex        =   20
         Top             =   1440
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   1085
         Caption         =   "export gradient file"
      End
      Begin PhotoDemon.pdButton cmdFile 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   1085
         Caption         =   "save to gradient collection"
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
      DontAutoReset   =   -1  'True
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   8220
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
      DontResetAutomatically=   -1  'True
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4575
      Index           =   1
      Left            =   0
      Top             =   1200
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   8070
      Begin PhotoDemon.pdPictureBoxInteractive picInteract 
         Height          =   330
         Left            =   0
         Top             =   2400
         Width           =   12615
         _ExtentX        =   0
         _ExtentY        =   0
      End
      Begin PhotoDemon.pdPictureBox picNodePreview 
         Height          =   1950
         Left            =   240
         Top             =   360
         Width           =   12135
         _ExtentX        =   0
         _ExtentY        =   0
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
         Width           =   12630
         _ExtentX        =   22278
         _ExtentY        =   503
         Alignment       =   2
         Caption         =   "yes"
         FontSize        =   9
         Layout          =   1
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4620
      Index           =   2
      Left            =   0
      Top             =   1200
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   8149
      Begin PhotoDemon.pdButton cmdRandomize 
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   3780
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1085
         Caption         =   "generate new pattern"
         FontSize        =   11
      End
      Begin PhotoDemon.pdPictureBox picAutoPreview 
         Height          =   1575
         Left            =   240
         Top             =   360
         Width           =   12135
         _ExtentX        =   0
         _ExtentY        =   0
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   1
         Left            =   120
         Top             =   0
         Width           =   9255
         _ExtentX        =   16536
         _ExtentY        =   556
         Caption         =   "preview"
         FontSize        =   12
      End
      Begin PhotoDemon.pdSlider sldOpacityAuto 
         Height          =   585
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   3000
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1032
         Max             =   100
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.pdColorSelector csColorAuto 
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1508
         Caption         =   "start color and opacity"
         curColor        =   0
      End
      Begin PhotoDemon.pdSlider sldOpacityAuto 
         Height          =   585
         Index           =   1
         Left            =   3720
         TabIndex        =   12
         Top             =   3000
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1032
         Max             =   100
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.pdColorSelector csColorAuto 
         Height          =   855
         Index           =   1
         Left            =   3600
         TabIndex        =   13
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1508
         Caption         =   "end color and opacity"
      End
      Begin PhotoDemon.pdSlider sldDensityAuto 
         Height          =   735
         Left            =   7200
         TabIndex        =   14
         Top             =   2040
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   1296
         Caption         =   "noise density"
         Max             =   100
         SigDigits       =   1
         Value           =   10
         NotchPosition   =   2
         NotchValueCustom=   10
      End
      Begin PhotoDemon.pdSlider sldVaryHSV 
         Height          =   735
         Index           =   0
         Left            =   7200
         TabIndex        =   16
         Top             =   2940
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1296
         Caption         =   "vary hue"
         FontSizeCaption =   10
         Max             =   100
         SigDigits       =   1
         ScaleStyle      =   1
      End
      Begin PhotoDemon.pdSlider sldVaryHSV 
         Height          =   735
         Index           =   1
         Left            =   9840
         TabIndex        =   17
         Top             =   2940
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1296
         Caption         =   "vary saturation"
         FontSizeCaption =   10
         Max             =   100
         SigDigits       =   1
         ScaleStyle      =   1
      End
      Begin PhotoDemon.pdSlider sldVaryHSV 
         Height          =   735
         Index           =   2
         Left            =   7200
         TabIndex        =   18
         Top             =   3720
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1296
         Caption         =   "vary luminance"
         FontSizeCaption =   10
         Max             =   100
         SigDigits       =   1
         ScaleStyle      =   1
         Value           =   20
         DefaultValue    =   20
      End
      Begin PhotoDemon.pdSlider sldVaryHSV 
         Height          =   735
         Index           =   3
         Left            =   9840
         TabIndex        =   19
         Top             =   3720
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1296
         Caption         =   "vary alpha"
         FontSizeCaption =   10
         Max             =   100
         SigDigits       =   1
         ScaleStyle      =   1
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6975
      Index           =   0
      Left            =   0
      Top             =   1200
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   12303
      Begin PhotoDemon.pdButton cmdEdit 
         Height          =   1575
         Left            =   9240
         TabIndex        =   23
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2778
         Caption         =   "edit this gradient >>"
         FontSize        =   12
      End
      Begin PhotoDemon.pdButtonStripVertical btsSort 
         Height          =   4335
         Left            =   9120
         TabIndex        =   22
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   10821
         Caption         =   "sort collection by"
      End
      Begin PhotoDemon.pdHyperlink lblCollection 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   6540
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   661
         Alignment       =   2
         Caption         =   ""
         RaiseClickEvent =   -1  'True
      End
      Begin PhotoDemon.pdListBoxOD lstGradients 
         Height          =   6375
         Left            =   240
         TabIndex        =   21
         Top             =   0
         Width           =   8655
         _ExtentX        =   15901
         _ExtentY        =   10821
         Caption         =   "your collection"
      End
   End
End
Attribute VB_Name = "dialog_GradientEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gradient Editor Dialog
'Copyright 2014-2026 by Tanner Helland
'Created: 23/July/15 (but assembled from many bits written earlier)
'Last updated: 13/April/22
'Last update: replace lingering picture boxes with pdPictureBox
'
'Comprehensive gradient editor.  This dialog gives the user multiple mechanisms for constructing,
' loading, and saving unique gradients.  It is used in multiple contexts, including the standalone
' gradient tool and gradient fill patterns.
'
'Note that - by design - this editor always returns a gradient with the same shape and angle as
' it was passed.  The editor itself doesn't handle gradient shape, angle, or other related properties,
' because we want gradient modifications to only happen in linear mode (matching the construction UI).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK/Cancel result from the dialog
Private m_userAnswer As VbMsgBoxResult

'The original gradient when the dialog was first loaded
Private m_OldGradient As String

'Gradient strings are generated with the help of PD's core gradient class.
' (NOTE: within this dialog, the gradient's shape is ignored; only a *linear* gradient is displayed, to make editing easier.)
Private m_NodePreview As pd2DGradient, m_AutoPreview As pd2DGradient

'If a user control spawned this dialog, it will pass itself as a reference.  We can then send gradient updates back
' to the control, allowing for real-time updates on the screen despite a modal dialog being raised!
Private parentGradientControl As pdGradientSelector

'Recently used gradients are loaded to/saved from a custom XML file
Private m_XMLEngine As pdXML

'The file where we'll store recent gradient data when the program is closed.  (At present, this file is located in PD's
' /Data/Presets/ folder.
Private m_XMLFilename As String

'Gradient preview DIB (required for color management) and interaction DIB (where all the gradient nodes are rendered)
Private m_NodePreviewDIB As pdDIB, m_InteractiveDIB As pdDIB, m_AutoPreviewDIB As pdDIB

'To prevent recursive setting changes, this value can be set to TRUE to prevent automatic UI synchronization
Private m_SuspendUI As Boolean

'This interface tracks its own collection of gradient points
Private m_NumOfGradientPoints As Long
Private m_GradientPoints() As GradientPoint

'In our assembled gradient collection, some items are just placeholders for special gradients -
' like "foreground to transparent", which are not hardcoded files but are assembled at run-time
' using the users current colors.
Private Enum PD_GradientSpecial
    gs_None = 0
    gs_FGtoBlack = 1
    gs_FGtoWhite = 2
    gs_FGtoTransparent = 3
End Enum

#If False Then
    Private Const gs_None = 0, gs_FGtoBlack = 1, gs_FGtoWhite = 2, gs_FGtoTransparent = 3
#End If

'The gradient collection is assembled at run-time from the files in the /Data/Gradients subfolder.
Private Type PD_GradientCollection
    gcPath As String
    gcFilename As String
    gcGradient As pd2DGradient
    gcThumb As pdDIB
    gcIsSpecial As PD_GradientSpecial
    gcGradientLoadedOK As Boolean
    gcLoadAttempted As Boolean
    gcDefaultIndex As Long
    gcAverageHue As Single
    gcAverageSaturation As Single
    gcAverageLuminance As Single
    gcAveragesCalculated As Boolean
End Type

Private m_GradientCollection() As PD_GradientCollection
Private m_NumGradientsInCollection As Long
Private Const GC_THUMB_WIDTH As Long = 200

Private Enum GC_SortOptions
    so_Filename = 0
    so_Name = 1
    so_Hue = 2
    so_Saturation = 3
    so_Luminance = 4
    so_Complexity = 5
End Enum

#If False Then
    Private Const so_Filename = 0, so_Name = 1, so_Hue = 2, so_Saturation = 3, so_Luminance = 4, so_Complexity = 5
#End If

'The user is allowed to sort their gradient collection by various criteria.  Because recursion is painfully slow in VB,
' our QuickSort implementation uses a stack instead.
Private Type QSStack
    sLB As Long
    sUB As Long
End Type

Private Const INIT_QUICKSORT_STACK_SIZE As Long = 512
Private m_qsStack() As QSStack
Private m_qsStackPtr As Long

'Height of the individual list items in the gradient collection preview list
Private Const BLOCKHEIGHT As Long = 52

'Font object(s) for rendering text in the individual list item boxes
Private m_ListFont As pdFont, m_ListFontTitle As pdFont

'Local list of themable colors.  Note that this is simply a duplicate of the color list in PD's Metadata Editor;
' the reason for this is simply to reuse the UI patterns of another (complicated) owner-drawn list box.
Private Enum GradientUI_ColorList
    [_First] = 0
    cl_TitleSelected = 0
    cl_TitleUnselected = 1
    cl_DescriptionSelected = 2
    cl_DescriptionUnselected = 3
    [_Last] = 3
    [_Count] = 4
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'The current gradient point (index) selected and/or hovered by the mouse.  -1 if no point is currently selected/hovered.
Private m_CurPoint As Long, m_CurHoverPoint As Long

'Similarly, the current mouse hover position over the interactive picture box. -1 if unhovered.
Private m_CurHoverX As Long

'Size of gradient "nodes" in the interactive UI.
Private Const GRADIENT_NODE_WIDTH As Single = 12!
Private Const GRADIENT_NODE_HEIGHT As Single = 14!

'Other gradient node UI renderers
Private m_inactiveArrowFill As pd2DBrush, m_activeArrowFill As pd2DBrush
Private m_inactiveOutlinePen As pd2DPen, m_activeOutlinePen As pd2DPen

'pdRandomize is used to create repeatable random patterns
Private m_Random As pdRandomize, m_RandomKey As Double

'When switching between panels, we need to know which panel was *previously* selected.
' This allows us to do things like auto-populate settings across panels.
' (Note that a separate value is used for the "reset" button on the command bar; because this
' is a separate control, we handle it as a special case, and restore the user's previous panel
' automatically.)
Private m_PreviousPanel As Long, m_PanelBeforeReset As Long, m_PrevPanelBeforeReset As Long

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_userAnswer
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
    If (btsEdit.ListIndex = 0) Then
        If (Not m_NodePreview Is Nothing) Then tmpGradient.CreateGradientFromString m_NodePreview.GetGradientAsString
    ElseIf (btsEdit.ListIndex = 1) Then
        If (Not m_NodePreview Is Nothing) Then tmpGradient.CreateGradientFromString m_NodePreview.GetGradientAsString
    ElseIf (btsEdit.ListIndex = 2) Then
        If (Not m_AutoPreview Is Nothing) Then tmpGradient.CreateGradientFromString m_AutoPreview.GetGradientAsString
    End If
    
    tmpGradient.SetGradientShape initGradient.GetGradientShape
    tmpGradient.SetGradientAngle initGradient.GetGradientAngle
    tmpGradient.SetGradientWrapMode initGradient.GetGradientWrapMode
    
    GetGradientAsOriginalShape = tmpGradient.GetGradientAsString()
    
End Function

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByVal initialGradient As String, Optional ByRef callingControl As pdGradientSelector = Nothing)
    
    'Store a reference to the calling control (if any)
    Set parentGradientControl = callingControl

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_userAnswer = vbCancel
    
    'Cache the initial gradient parameters so we can access it elsewhere
    m_OldGradient = initialGradient
    
    'Inside this dialog, the gradient is always forced to a linear-type gradient at angle 0.  This makes it much easier to edit.
    Set m_NodePreview = New pd2DGradient
    m_NodePreview.CreateGradientFromString m_OldGradient
    m_NodePreview.SetGradientShape P2_GS_Linear
    m_NodePreview.SetGradientAngle 0#
    
    'If the dialog is being initialized for the first time, there will be no "initial gradient".  In this case, the gradient class
    ' will initialize a placeholder gradient.  We make a copy of it, and use that as the basis of the editor's initial settings.
    If (LenB(m_OldGradient) = 0) Then m_OldGradient = m_NodePreview.GetGradientAsString
    
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
    m_XMLFilename = UserPrefs.GetPresetPath & "Gradient_Selector.xml"
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub btsEdit_Click(ByVal buttonIndex As Long)
    ChangeActivePanel buttonIndex
End Sub

Private Sub ChangeActivePanel(ByVal buttonIndex As Long)
    
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        
        'Show the primary container for this panel
        picContainer(i).Visible = (i = buttonIndex)
        
        'Show/hide the import/export buttons as necessary - they are only relevant for the last two panels
        pnlShared.Visible = (buttonIndex <> 0)
        
        'Noise gradients don't support "distribute points evenly", as it makes no sense there
        chkDistributeEvenly.Visible = (buttonIndex = 1)
        
    Next i
    
    'When changing from "manual" to "auto" mode, mirror the "manual" mode settings to the "automatic" pane.
    On Error GoTo ContinueWithPreview
    
    If (buttonIndex = 2) And (m_PreviousPanel = 1) Then
        
        If (m_NumOfGradientPoints > 1) Then
        
            'Find the left-most and right-most colors in the current gradient arrangement
            Dim minPos As Single, minPosIndex As Long, maxPos As Single, maxPosIndex As Long
            minPos = 1#
            maxPos = 0#
            
            For i = 0 To m_NumOfGradientPoints - 1
                If (m_GradientPoints(i).PointPosition < minPos) Then
                    minPos = m_GradientPoints(i).PointPosition
                    minPosIndex = i
                ElseIf (m_GradientPoints(i).PointPosition > maxPos) Then
                    maxPos = m_GradientPoints(i).PointPosition
                    maxPosIndex = i
                End If
            Next i
            
            csColorAuto(0).Color = m_GradientPoints(minPosIndex).PointRGB
            sldOpacityAuto(0).Value = m_GradientPoints(minPosIndex).PointOpacity
            csColorAuto(1).Color = m_GradientPoints(maxPosIndex).PointRGB
            sldOpacityAuto(1).Value = m_GradientPoints(maxPosIndex).PointOpacity
            
        End If
        
    End If
        
ContinueWithPreview:
    
    UpdatePreview
    m_PreviousPanel = buttonIndex
    
End Sub

Private Sub btsSort_Click(ByVal buttonIndex As Long)
    ChangeCollectionOrder buttonIndex
End Sub

'Change the listed order of the gradient collection.  This requires sorting against different sort criteria.
Private Sub ChangeCollectionOrder(ByVal sortCriteria As GC_SortOptions)

    If (m_NumGradientsInCollection <= 1) Then Exit Sub
    
    Dim i As Long
    
    Select Case sortCriteria
    
        'Default sort order
        Case so_Filename
        
            'Default sort order simply uses filenames.  This is nice as we don't have to actually load
            ' individual gradient files to have sortable data.
            
        'Gradient name
        Case so_Name
        
            'Gradients are not loaded by default at run-time; instead, we load them on-demand.  Sorting by
            ' name requires us to load each gradient, so pre-load them now.
            LoadAllGradients
        
        'H/S/L
        Case so_Hue, so_Saturation, so_Luminance
        
            'Gradients must all be loaded, and individual gradient classes must be queried for average
            ' HSL values.  (Once calculated, however, all three HSL values are cached locally, making
            ' subsequent sorts very fast.)
            LoadAllGradients
            
            For i = 0 To m_NumGradientsInCollection - 1
                With m_GradientCollection(i)
                    If (Not .gcAveragesCalculated) Then
                        .gcGradient.GetAverageHSL .gcAverageHue, .gcAverageSaturation, .gcAverageLuminance
                        .gcAveragesCalculated = True
                    End If
                End With
            Next i
            
        'Complexity (e.g. number of stops)
        Case so_Complexity
            LoadAllGradients
    
    End Select
    
    'After any prerequisites are filled, perform the sort
    QuickSortStringStack sortCriteria
    
    'Repopulate the listbox
    lstGradients.SetAutomaticRedraws True, True
    
End Sub

'Some sort criteria require us to load all gradients.  Call this function to do so (and note that it's harmless
' to call, in general, as already loaded gradients will not be re-loaded).
Private Sub LoadAllGradients()
    Dim i As Long
    For i = 0 To m_NumGradientsInCollection - 1
        LoadGradientCollectionPreview i
    Next i
End Sub

Private Sub QuickSortStringStack(ByVal sortCriteria As GC_SortOptions)
    
    If (m_NumGradientsInCollection > 1) Then
    
        'Prep our internal stack
        ReDim m_qsStack(0 To INIT_QUICKSORT_STACK_SIZE - 1) As QSStack
        m_qsStackPtr = 0
        m_qsStack(0).sLB = 0
        m_qsStack(0).sUB = m_NumGradientsInCollection - 1
        
        NaiveQuickSortExtended sortCriteria
        
        'Free the stack before exiting
        Erase m_qsStack
        
    End If
    
End Sub

'Semi-standard QuickSort implementation, with VB-specific enhancements provided by georgekar, and further
' enhancements by myself to further improve performance.
'
'georgekar's original, unmodified implementation can be found here:
' http://www.vbforums.com/showthread.php?781043-VB6-Dual-Pivot-QuickSort
Private Sub NaiveQuickSortExtended(ByVal sortCriteria As Long)
    
    Dim lowVal As Long, highVal As Long
    Dim i As Long, j As Long, v As PD_GradientCollection
    
    Do
        
        'Load the next set of boundaries, and reset all pivots
        lowVal = m_qsStack(m_qsStackPtr).sLB
        highVal = m_qsStack(m_qsStackPtr).sUB
        
        'Check for single-entry ranges
        If (highVal - lowVal = 1) Then
            i = lowVal
            If (CompareIndices(i, highVal, sortCriteria) > 0) Then SwapIndices i, highVal
            GoTo NextSortItem
        Else
            
            'Bisect this range
            i = (lowVal + highVal) \ 2
            
            'Migrate all equal entries into place
            If (CompareIndices(i, lowVal, sortCriteria) = 0) Then
                
                j = highVal - 1
                i = lowVal
                
                Do
                    i = i + 1
                    If (i > j) Then
                        If (CompareIndices(highVal, lowVal, sortCriteria) < 0) Then SwapIndices lowVal, highVal
                        GoTo NextSortItem
                    End If
                Loop Until (CompareIndices(i, lowVal, sortCriteria) <> 0)
                
                v = m_GradientCollection(i)
                If (i > lowVal) Then If (CompareIndices(lowVal, i, sortCriteria) > 0) Then SwapIndices lowVal, i
                
            'Move the pointer until we arrive at an unsorted pivot
            Else
                v = m_GradientCollection(i)
                i = lowVal
                Do While (CompareValues(m_GradientCollection(i), v, sortCriteria) < 0): i = i + 1: Loop
            End If

        'End special case handling
        End If
        
        'Resume standard QuickSort behavior
        j = highVal
        
        Do
            'Advance from the right
            Do While (CompareValues(m_GradientCollection(j), v, sortCriteria) > 0): j = j - 1: Loop
            
            'Swap as necessary
            If (i <= j) Then
                SwapIndices i, j
                i = i + 1
                j = j - 1
            End If
            
            If (i > j) Then Exit Do
            
            'Advance from the left
            Do While (CompareValues(m_GradientCollection(i), v, sortCriteria) < 0): i = i + 1: Loop
            
        Loop
        
        'Conditionally add new entries to the processing stack
        If (lowVal < j) Then
            m_qsStack(m_qsStackPtr).sLB = lowVal
            m_qsStack(m_qsStackPtr).sUB = j
            m_qsStackPtr = m_qsStackPtr + 1
        End If
        
        If (i < highVal) Then
            m_qsStack(m_qsStackPtr).sLB = i
            m_qsStack(m_qsStackPtr).sUB = highVal
            m_qsStackPtr = m_qsStackPtr + 1
        End If
        
'Yep, VB6 requires us to use GOTO and line labels.  There is no "Continue For" equivalent.
NextSortItem:
        
        'Decrement the stack pointer
        m_qsStackPtr = m_qsStackPtr - 1
        
    Loop While (m_qsStackPtr >= 0)
    
End Sub

'Helper for gradient sort functions.  Returns -1 for "<", 0 for "=", and 1 for ">".
Private Function CompareIndices(ByVal idx1 As Long, ByVal idx2 As Long, ByVal sortCriteria As GC_SortOptions) As Long
    CompareIndices = CompareValues(m_GradientCollection(idx1), m_GradientCollection(idx2), sortCriteria)
End Function

'Helper for gradient sort functions.  Returns -1 for "<", 0 for "=", and 1 for ">".
Private Function CompareValues(ByRef tmpGradientEntry1 As PD_GradientCollection, ByRef tmpGradientEntry2 As PD_GradientCollection, ByVal sortCriteria As GC_SortOptions) As Long
    
    Select Case sortCriteria
    
        'Gradient filename
        Case so_Filename
            CompareValues = Strings.StrCompSort(tmpGradientEntry1.gcFilename, tmpGradientEntry2.gcFilename)
            
        'Gradient name (requires gradient files to be loaded)
        Case so_Name
            CompareValues = Strings.StrCompSort(tmpGradientEntry1.gcGradient.GetGradientName, tmpGradientEntry2.gcGradient.GetGradientName)
            
        'H/S/L
        Case so_Hue
            If (tmpGradientEntry1.gcAverageHue < tmpGradientEntry2.gcAverageHue) Then
                CompareValues = -1
            ElseIf (tmpGradientEntry1.gcAverageHue > tmpGradientEntry2.gcAverageHue) Then
                CompareValues = 1
            Else
                CompareValues = 0
            End If
            
        Case so_Saturation
            If (tmpGradientEntry1.gcAverageSaturation < tmpGradientEntry2.gcAverageSaturation) Then
                CompareValues = -1
            ElseIf (tmpGradientEntry1.gcAverageSaturation > tmpGradientEntry2.gcAverageSaturation) Then
                CompareValues = 1
            Else
                CompareValues = 0
            End If
            
        Case so_Luminance
            If (tmpGradientEntry1.gcAverageLuminance < tmpGradientEntry2.gcAverageLuminance) Then
                CompareValues = -1
            ElseIf (tmpGradientEntry1.gcAverageLuminance > tmpGradientEntry2.gcAverageLuminance) Then
                CompareValues = 1
            Else
                CompareValues = 0
            End If
        
        'Complexity (no. of stops)
        Case so_Complexity
            If (tmpGradientEntry1.gcGradient.GetNumOfNodes < tmpGradientEntry2.gcGradient.GetNumOfNodes) Then
                CompareValues = -1
            ElseIf (tmpGradientEntry1.gcGradient.GetNumOfNodes > tmpGradientEntry2.gcGradient.GetNumOfNodes) Then
                CompareValues = 1
            Else
                CompareValues = 0
            End If
    
    End Select
    
End Function

'Helper for gradient sort functions.  Feel free to modify to improve swap performance, depending on the data type being sorted.
Private Sub SwapIndices(ByVal objIndex1 As Long, ByVal objIndex2 As Long)
    Dim tmpGradientCollection As PD_GradientCollection
    tmpGradientCollection = m_GradientCollection(objIndex1)
    m_GradientCollection(objIndex1) = m_GradientCollection(objIndex2)
    m_GradientCollection(objIndex2) = tmpGradientCollection
End Sub

Private Sub chkDistributeEvenly_Click()
    If (Not m_SuspendUI) Then
        RedrawEverything
        SyncUIToActiveNode
    End If
End Sub

Private Sub chkGamma_Click()
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

Private Sub cmdBar_BeforeResetClick(ByRef cancelReset As Boolean)
    m_PrevPanelBeforeReset = m_PreviousPanel
    m_PanelBeforeReset = btsEdit.ListIndex
End Sub

'CANCEL BUTTON
Private Sub cmdBar_CancelClick()
    m_userAnswer = vbCancel
    Me.Hide
End Sub

'OK BUTTON
Private Sub cmdBar_OKClick()

    'Mirror the selected gradient's settings to the node preview gradient (which controls what gradient
    ' data is returned to the caller.)
    If (btsEdit.ListIndex = 0) Then
        If (lstGradients.ListIndex >= 0) Then
            Set m_NodePreview = m_GradientCollection(lstGradients.ListIndex).gcGradient
            With m_NodePreview
                .SetGradientShape P2_GS_Linear
                .SetGradientAngle 0#
            End With
        End If
    
    'Store the newGradient value (which the dialog handler will use to return the selected gradient to the caller)
    Else
        UpdateGradientObjects
    End If
    
    m_userAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_RandomizeClick()
    
    btsEdit.ListIndex = m_PreviousPanel
    txtName.Text = vbNullString
    
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_AutomaticAndRandom
    
    'Create several random nodes with random colors and opacities
    m_NumOfGradientPoints = 2 + Int(cRandom.GetRandomFloat_WH * 4#)
    ReDim m_GradientPoints(0 To m_NumOfGradientPoints - 1) As GradientPoint
    
    Dim i As Long
    For i = 0 To m_NumOfGradientPoints - 1
        With m_GradientPoints(i)
            .PointOpacity = 100#
            .PointPosition = cRandom.GetRandomFloat_WH
            .PointRGB = cRandom.GetRandomFloat_WH * &HFFFFFF
        End With
    Next i
    
    RedrawEverything
    
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
    
    m_SuspendUI = True
    
    m_PreviousPanel = m_PrevPanelBeforeReset
    btsEdit.ListIndex = m_PanelBeforeReset
    
    'Reset our central gradient object; everything else derives from it
    Set m_NodePreview = New pd2DGradient
    m_NodePreview.CreateGradientFromString vbNullString
    
    m_NumOfGradientPoints = 2
    ReDim m_GradientPoints(0 To 1) As GradientPoint
    
    m_GradientPoints(0).PointOpacity = 100#
    m_GradientPoints(0).PointPosition = 0#
    m_GradientPoints(0).PointRGB = vbBlack
    
    m_GradientPoints(1).PointOpacity = 100#
    m_GradientPoints(1).PointPosition = 1#
    m_GradientPoints(1).PointRGB = vbWhite
    
    Me.csColorAuto(0).Color = vbBlack
    
    chkDistributeEvenly.Value = False
    chkGamma.Value = False
    txtName.Text = vbNullString
    
    m_SuspendUI = False
    
    'Synchronize all controls to the updated settings
    UpdateGradientObjects
    SyncControlsToGradientObject
    UpdatePreview
    
End Sub

Private Sub cmdEdit_Click()

    If (lstGradients.ListIndex >= 0) Then
        Set m_NodePreview = m_GradientCollection(lstGradients.ListIndex).gcGradient
        
        'Sync all controls to reflect the new gradient
        m_CurPoint = -1
        chkDistributeEvenly.Value = False
        SyncControlsToGradientObject
        
        UpdatePreview
            
        'Switch the active panel
        btsEdit.ListIndex = 1
        
    End If

End Sub

Private Sub cmdFile_Click(Index As Integer)

    Select Case Index
        
        'Save to your collection
        Case 0
            
            'Start by making sure the current name exists and is unique
            If (LenB(txtName.Text) > 0) Then
            
                'Name exists; see if it is unique in the current collection.  (Unfortunately, this requires us
                ' to manually load all gradient files.)
                LoadAllGradients
                
                Dim nameIsUnique As Boolean
                nameIsUnique = True
                
                Dim i As Long, matchingIndex As Long
                matchingIndex = -1
                
                For i = 0 To m_NumGradientsInCollection - 1
                    If Strings.StringsEqual(m_GradientCollection(i).gcGradient.GetGradientName, txtName.Text, True) Then
                        nameIsUnique = False
                        matchingIndex = i
                        Exit For
                    End If
                Next i
                
                'If the name is unique, save it directly to the user's gradient folder.
                Dim dstFilename As String
                
                If nameIsUnique Then
                    dstFilename = Files.IncrementFilename(UserPrefs.GetGradientPath(), txtName.Text, "svg") & ".svg"
                    
                'If the name is *not* unique, provide an overwrite prompt.
                Else
                    
                    Dim pResult As VbMsgBoxResult
                    pResult = PDMsgBox("A gradient with the name ""%1"" already exists in your collection.  Would you like to overwrite it?", vbYesNoCancel Or vbExclamation, "Duplicate name", txtName.Text)
                    If (pResult = vbNo) Or (pResult = vbCancel) Then Exit Sub
                    dstFilename = m_GradientCollection(i).gcFilename
                    
                End If
                
                'Kill the existing file, if any
                dstFilename = UserPrefs.GetGradientPath() & dstFilename
                Files.FileDeleteIfExists dstFilename
                
                'Make sure all gradient settings are up-to-date
                UpdateGradientObjects
                
                'Save based on the provided extension (which allows us to overwrite ggr files with new ggr data, as relevant)
                If Strings.StringsEqual(Files.FileGetExtension(dstFilename), "ggr", True) Then
                    If (btsEdit.ListIndex = 1) Then m_NodePreview.SaveGradient_GIMP dstFilename Else m_AutoPreview.SaveGradient_GIMP dstFilename
                Else
                    If (btsEdit.ListIndex = 1) Then m_NodePreview.SaveGradient_SVG dstFilename Else m_AutoPreview.SaveGradient_SVG dstFilename
                End If
                
                'Reset the gradient collection so that the new addition is picked-up
                BuildGradientCollection
                
                'Automatically select the just-saved gradient in the gradient list
                dstFilename = Files.FileGetName(dstFilename)
                For i = 0 To m_NumGradientsInCollection - 1
                    If Strings.StringsEqual(m_GradientCollection(i).gcFilename, dstFilename, True) Then
                        lstGradients.ListIndex = i
                        Exit For
                    End If
                Next i
            
            'Name does not exist; ask the user to supply one.
            Else
                PDMsgBox "This gradient doesn't have a name.  Please give it a name before adding it to your collection.", vbOKOnly Or vbInformation, "Name required"
            End If
        
        'Load gradient file
        Case 1
            ImportGradientFile
        
        'Save gradient file
        Case 2
            ExportGradientFile
    
    End Select
    
End Sub

Private Sub ImportGradientFile()

    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    'Build a common dialog filter list
    Dim cdFilter As pdString
    Set cdFilter = New pdString
    
    cdFilter.Append g_Language.TranslateMessage("All supported gradients") & "|*.svg;*.ggr|"
    cdFilter.Append g_Language.TranslateMessage("GIMP Gradient") & " (.ggr)|*.ggr|"
    cdFilter.Append g_Language.TranslateMessage("SVG Gradient") & " (.svg)|*.svg|"
    cdFilter.Append g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdIndex As Long
    cdIndex = 1
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Import gradient")
    
    'Prep a common dialog interface
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    Dim srcFilename As String
    If openDialog.GetOpenFileName(srcFilename, , True, False, cdFilter.ToString(), cdIndex, UserPrefs.GetGradientPath(), cdTitle, , GetModalOwner().hWnd) Then
        
        'Update preferences
        UserPrefs.SetGradientPath Files.FileGetPath(srcFilename)
        
        'For now, forcibly switch to the "manual" panel and load the gradient there
        If (btsEdit.ListIndex <> 1) Then btsEdit.ListIndex = 1
        
        Dim tmpGradient As pd2DGradient
        Set tmpGradient = New pd2DGradient
        tmpGradient.SetGradientAngle 0!
        tmpGradient.SetGradientShape P2_GS_Linear
        
        If tmpGradient.LoadGradientFromFile(srcFilename) Then
            
            Set m_NodePreview = tmpGradient
            
            'Sync all controls to reflect the new gradient
            m_CurPoint = -1
            chkDistributeEvenly.Value = False
            SyncControlsToGradientObject
            
            UpdatePreview
            
        Else
            PDMsgBox "Unfortunately, the gradient file ""%1"" doesn't appear to be a valid gradient file.", vbInformation Or vbOKOnly Or vbApplicationModal, "Invalid gradient"
        End If
        
    End If
    
    'Re-enable UI
    Interface.EnableUserInput
    
End Sub

Private Sub ExportGradientFile()

    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    'Determine an initial folder.  This is easy - just grab the last "profile" path from the preferences file.
    Dim initialSaveFolder As String
    initialSaveFolder = UserPrefs.GetGradientPath()
    
    'Build a common dialog filter list
    Dim cdFilter As pdString, cdFilterExtensions As pdString
    Set cdFilter = New pdString
    Set cdFilterExtensions = New pdString
    
    cdFilter.Append g_Language.TranslateMessage("GIMP Gradient") & " (.ggr)|*.ggr|"
    cdFilter.Append g_Language.TranslateMessage("SVG Gradient") & " (.svg)|*.svg"
    cdFilterExtensions.Append "ggr|"
    cdFilterExtensions.Append "svg"
    
    Dim cdIndex As Long
    cdIndex = 2
    
    'Suggest a file name.  If the user has entered a gradient name, we suggest that first.
    ' After that, we attempt to reuse the current image's name (if any).
    Dim dstFilename As String
    If (LenB(txtName.Text) > 0) Then
        dstFilename = txtName.Text
    ElseIf (PDImages.GetNumOpenImages > 0) Then
        dstFilename = PDImages.GetActiveImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    End If
    
    'If none of our previous name suggestions stuck, suggest a default name.
    If (LenB(dstFilename) = 0) Then dstFilename = g_Language.TranslateMessage("New gradient")
    dstFilename = initialSaveFolder & dstFilename
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Export gradient")
    
    'Display a common save dialog
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    If saveDialog.GetSaveFileName(dstFilename, , True, cdFilter.ToString(), cdIndex, initialSaveFolder, cdTitle, cdFilterExtensions.ToString(), GetModalOwner().hWnd) Then
    
        'Update preferences
        UserPrefs.SetGradientPath Files.FileGetPath(dstFilename)
        
        'Set the source gradient differently, depending on the current active panel
        Dim srcGradient As pd2DGradient
        If (btsEdit.ListIndex = 1) Then
            Set srcGradient = m_NodePreview
        ElseIf (btsEdit.ListIndex = 2) Then
            Set srcGradient = m_AutoPreview
        End If
                
        'Proceed with saving.  The embedded gradient class handles this for us.
        Select Case cdIndex
            
            'Export GIMP format
            Case 1
                srcGradient.SaveGradient_GIMP dstFilename
                
            'Export SVG format
            Case 2
                srcGradient.SaveGradient_SVG dstFilename
            
            'No other supported formats at present
            Case Else
            
        End Select
        
    End If
    
    'Re-enable UI
    Interface.EnableUserInput
        
End Sub

Private Sub cmdRandomize_Click()
    If (m_Random Is Nothing) Then Set m_Random = New pdRandomize
    m_Random.SetSeed_AutomaticAndRandom
    m_RandomKey = m_Random.GetSeed()
    UpdatePreview
End Sub

Private Sub csColorAuto_ColorChanged(Index As Integer)
    UpdatePreview
End Sub

Private Sub csNode_ColorChanged()
    
    If (Not m_SuspendUI) And (m_CurPoint >= 0) Then
        m_GradientPoints(m_CurPoint).PointRGB = csNode.Color
        RedrawEverything
    End If
    
End Sub

Private Sub Form_Load()
    
    'Add the instructions label
    Dim instructionText As String
    instructionText = g_Language.TranslateMessage("Left-click to add new nodes or edit existing nodes.  Right-click a node to remove it.")
    lblInstructions.Caption = instructionText
    
    'Populate button strips, drop-downs, tooltips, etc
    btsEdit.AddItem "gradient collection", 0
    btsEdit.AddItem "make your own", 1
    btsEdit.AddItem "noise gradients", 2
    btsEdit.ListIndex = 0
    m_PreviousPanel = 0
    
    btsSort.AddItem "filename", 0
    btsSort.AddItem "gradient name", 1
    btsSort.AddItem "hue", 2
    btsSort.AddItem "saturation", 3
    btsSort.AddItem "luminance", 4
    btsSort.AddItem "complexity", 5
    btsSort.ListIndex = 0
    
    lstGradients.ListItemHeight = Interface.FixDPI(BLOCKHEIGHT)
    
    Dim buttonImgSize As Long
    buttonImgSize = Interface.FixDPI(24)
    cmdFile(0).AssignImage "file_save", , buttonImgSize, buttonImgSize
    cmdFile(1).AssignImage "file_open", , buttonImgSize, buttonImgSize
    cmdFile(2).AssignImage "file_saveas", , buttonImgSize, buttonImgSize
    
    lblCollection.Caption = g_Language.TranslateMessage("Your gradient collection is stored in the ""%1"" folder", UserPrefs.GetGradientPath(True))
    lblCollection.AssignTooltip "click to open this folder in Windows Explorer"
    
    chkGamma.AssignTooltip "When a gradient contains colors with wildly different luminance values, gamma correction may improve its appearance."
    chkDistributeEvenly.AssignTooltip "Use this setting to automatically calculate equal positioning for all gradient nodes."
    
    If PDMain.IsProgramRunning() Then
    
        If (m_NodePreview Is Nothing) Then Set m_NodePreview = New pd2DGradient
        
        'Prep a default set of gradient points in the editor panel
        ResetGradientPoints
        
        'Load the gradient collection, including any associated rendering items
        Set m_Colors = New pdThemeColors
        Dim colorCount As GradientUI_ColorList: colorCount = [_Count]
        m_Colors.InitializeColorList "PDMetadataList", colorCount
        UpdateColorList
        
        Set m_ListFontTitle = New pdFont
        m_ListFontTitle.SetFontBold True
        m_ListFontTitle.SetFontSize 10
        m_ListFontTitle.CreateFontObject
        m_ListFontTitle.SetTextAlignment vbLeftJustify
        
        Set m_ListFont = New pdFont
        m_ListFont.SetFontBold False
        m_ListFont.SetFontSize 10
        m_ListFont.CreateFontObject
        m_ListFont.SetTextAlignment vbLeftJustify
        
        BuildGradientCollection
        cmdEdit.Enabled = (lstGradients.ListIndex >= 0)
        
        'Prep all gradient point tracking variables
        m_CurPoint = -1
        
        'While we're here, we'll also prep all generic drawing objects for the interactive gradient node UI bits
        Set m_inactiveArrowFill = New pd2DBrush
        Set m_activeArrowFill = New pd2DBrush
        
        m_inactiveArrowFill.SetBrushMode P2_BM_Solid
        m_inactiveArrowFill.SetBrushOpacity 100!
        m_inactiveArrowFill.SetBrushColor g_Themer.GetGenericUIColor(UI_Background)
        m_inactiveArrowFill.CreateBrush
        
        m_activeArrowFill.SetBrushMode P2_BM_Solid
        m_activeArrowFill.SetBrushOpacity 100!
        m_activeArrowFill.SetBrushColor g_Themer.GetGenericUIColor(UI_AccentLight)
        m_activeArrowFill.CreateBrush
        
        Set m_inactiveOutlinePen = New pd2DPen
        Set m_activeOutlinePen = New pd2DPen
        
        m_inactiveOutlinePen.SetPenStyle GP_DS_Solid
        m_inactiveOutlinePen.SetPenOpacity 100
        m_inactiveOutlinePen.SetPenWidth 1#
        m_inactiveOutlinePen.SetPenLineJoin GP_LJ_Miter
        m_inactiveOutlinePen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayDark)
        m_inactiveOutlinePen.CreatePen
        
        m_activeOutlinePen.SetPenStyle GP_DS_Solid
        m_activeOutlinePen.SetPenOpacity 100
        m_activeOutlinePen.SetPenWidth 1#
        m_activeOutlinePen.SetPenLineJoin GP_LJ_Miter
        m_activeOutlinePen.SetPenColor g_Themer.GetGenericUIColor(UI_Accent)
        m_activeOutlinePen.CreatePen
        
        'Draw the initial set of interactive gradient nodes
        SyncUIToActiveNode
        DrawGradientNodes
                
    End If
    
    ChangeActivePanel btsEdit.ListIndex
    RedrawEverything

End Sub

Private Sub BuildGradientCollection()

    'Disable automatic list box redraws (to improve performance when adding/removing large item counts)
    lstGradients.SetAutomaticRedraws False, False
    lstGradients.Clear
    
    'A pdFSO object will help us quickly iterate (potentially) valid gradient files
    Dim cFSO As pdFSO
    Set cFSO = New pdFSO
    
    'Always add three default gradients to the collection
    Const INIT_COLLECTION_SIZE As Long = 16
    ReDim m_GradientCollection(0 To INIT_COLLECTION_SIZE - 1) As PD_GradientCollection
    m_NumGradientsInCollection = 0
        
    Dim i As Long
    For i = 0 To 2
        With m_GradientCollection(i)
            .gcDefaultIndex = i
            Set .gcGradient = New pd2DGradient
            .gcGradientLoadedOK = True
            .gcLoadAttempted = True
            .gcIsSpecial = i + 1
            .gcFilename = "_"       'Ensure these gradients appear near the top of the "Default" collection sort
            Select Case .gcIsSpecial
                Case gs_FGtoBlack
                    .gcGradient.CreateTwoPointGradient layerpanel_Colors.GetCurrentColor(), RGB(0, 0, 0)
                    .gcGradient.SetGradientName g_Language.TranslateMessage("Foreground to black")
                Case gs_FGtoWhite
                    .gcGradient.CreateTwoPointGradient layerpanel_Colors.GetCurrentColor(), RGB(255, 255, 255)
                    .gcGradient.SetGradientName g_Language.TranslateMessage("Foreground to white")
                Case gs_FGtoTransparent
                    .gcGradient.CreateTwoPointGradient layerpanel_Colors.GetCurrentColor(), RGB(0, 0, 0), 100!, 0!
                    .gcGradient.SetGradientName g_Language.TranslateMessage("Foreground to transparent")
            End Select
        End With
        m_NumGradientsInCollection = m_NumGradientsInCollection + 1
    Next i
    
    Dim srcFiles As pdStringStack
    If cFSO.RetrieveAllFiles(UserPrefs.GetGradientPath(True), srcFiles, True, True, "ggr|svg") Then
        
        Dim tmpString As String
        For i = 0 To srcFiles.GetNumOfStrings - 1
            tmpString = srcFiles.GetString(i)
            If (m_NumGradientsInCollection > UBound(m_GradientCollection)) Then ReDim Preserve m_GradientCollection(0 To m_NumGradientsInCollection * 2 - 1) As PD_GradientCollection
            With m_GradientCollection(m_NumGradientsInCollection)
                .gcPath = tmpString
                .gcFilename = Files.FileGetName(tmpString)
                .gcDefaultIndex = m_NumGradientsInCollection
            End With
            m_NumGradientsInCollection = m_NumGradientsInCollection + 1
        Next i
        
    End If
    
    'Add all list items to the list box
    For i = 0 To m_NumGradientsInCollection - 1
        lstGradients.AddItem vbNullString, i
    Next i
    
    'After adding all files, perform a default sort by filename
    ChangeCollectionOrder so_Filename
    
    'Render the finished list
    lstGradients.SetAutomaticRedraws True, True
    
End Sub

Private Sub ResetGradientPoints()
    
    m_NumOfGradientPoints = 2
    ReDim m_GradientPoints(0 To m_NumOfGradientPoints - 1) As GradientPoint
    
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
    If chkDistributeEvenly.Value Then
        
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
            m_CurHoverX = -1
            
            SyncUIToActiveNode
            
            m_NodePreview.CreateGradientFromPointCollection m_NumOfGradientPoints, m_GradientPoints
            m_NodePreview.GetCopyOfPointCollection m_NumOfGradientPoints, m_GradientPoints
            
        End If
        
        'Redistribute points accordingly
        For i = 0 To m_NumOfGradientPoints - 1
            m_GradientPoints(i).PointPosition = i / (m_NumOfGradientPoints - 1)
        Next i
        
    End If
    
    'Manual edit mode...
    If (btsEdit.ListIndex = 1) Then
        With m_NodePreview
            .SetGradientShape P2_GS_Linear
            .SetGradientAngle 0#
            .CreateGradientFromPointCollection m_NumOfGradientPoints, m_GradientPoints
            .SetGradientGammaMode chkGamma.Value
            .SetGradientName txtName.Text
        End With
    
    'Noise gradient mode...
    Else
        If (Not m_AutoPreview Is Nothing) Then
            With m_AutoPreview
                .SetGradientShape P2_GS_Linear
                .SetGradientAngle 0#
                .SetGradientGammaMode chkGamma.Value
                .SetGradientName txtName.Text
            End With
        End If
    End If

End Sub

'Make all UI elements reflect the current gradient object.  This is typically done after the dialog loads, or after loading a
' previously created gradient.
Private Sub SyncControlsToGradientObject()
        
    m_SuspendUI = True
    
    With m_NodePreview
        .GetCopyOfPointCollection m_NumOfGradientPoints, m_GradientPoints
        chkGamma.Value = .GetGradientGammaMode()
        txtName.Text = .GetGradientName()
    End With
    
    DrawGradientNodes
    
    m_SuspendUI = False
    
    'Also, synchronize the node-specific UI to the active node (if any)
    SyncUIToActiveNode
    
End Sub

Private Sub lblCollection_Click()
    Dim filePath As String, shellCommand As String
    filePath = UserPrefs.GetGradientPath(True)
    shellCommand = "explorer.exe """ & filePath & """"
    Shell shellCommand, vbNormalFocus
End Sub

Private Sub lstGradients_Click()
    cmdEdit.Enabled = (lstGradients.ListIndex >= 0)
End Sub

Private Sub lstGradients_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)

    If (bufferDC = 0) Then Exit Sub
    
    'Start by loading the requested gradient if we haven't already.
    If (m_GradientCollection(itemIndex).gcGradient Is Nothing) Then LoadGradientCollectionPreview itemIndex
    
    '...Followed by a matching preview DIB
    If (m_GradientCollection(itemIndex).gcThumb Is Nothing) Then LoadGradientCollectionPreviewDIB itemIndex
    
    'Calculate text colors (which vary depending on selection state)
    Dim txtTitleColor As Long, txtDescriptionColor As Long
    If itemIsSelected Then
        txtTitleColor = m_Colors.RetrieveColor(cl_TitleSelected, lstGradients.Enabled, , itemIsHovered)
        txtDescriptionColor = m_Colors.RetrieveColor(cl_DescriptionSelected, lstGradients.Enabled, , itemIsHovered)
    Else
        txtTitleColor = m_Colors.RetrieveColor(cl_TitleUnselected, lstGradients.Enabled, , itemIsHovered)
        txtDescriptionColor = m_Colors.RetrieveColor(cl_DescriptionUnselected, lstGradients.Enabled, , itemIsHovered)
    End If
    
    'Retrieve the item's boundary rect
    Dim tmpRectF As RectF
    CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&
    
    Dim offsetX As Single, offsetY As Single
    offsetX = tmpRectF.Left + Interface.FixDPI(8)
    offsetY = tmpRectF.Top + (tmpRectF.Height - m_GradientCollection(itemIndex).gcThumb.GetDIBHeight) \ 2
    
    'Render the gradient preview first, with a light border around it.
    If m_GradientCollection(itemIndex).gcGradientLoadedOK Then m_GradientCollection(itemIndex).gcThumb.AlphaBlendToDC bufferDC, , offsetX, offsetY
    
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC
    
    Dim cPen As pd2DPen, penColor As Long
    If itemIsSelected Then
        penColor = g_Themer.GetGenericUIColor(UI_Accent, lstGradients.Enabled, , itemIsHovered)
    Else
        penColor = g_Themer.GetGenericUIColor(UI_GrayNeutral, lstGradients.Enabled, , itemIsHovered)
    End If
    Drawing2D.QuickCreateSolidPen cPen, 1!, penColor
    
    PD2D.DrawRectangleF cSurface, cPen, offsetX, offsetY, m_GradientCollection(itemIndex).gcThumb.GetDIBWidth, m_GradientCollection(itemIndex).gcThumb.GetDIBHeight
    
    Set cSurface = Nothing
    Set cPen = Nothing
    
    'Next, render the gradient's name and path
    offsetX = offsetX + m_GradientCollection(itemIndex).gcThumb.GetDIBWidth + Interface.FixDPI(8)
    offsetY = offsetY + 2
    
    If (LenB(m_GradientCollection(itemIndex).gcGradient.GetGradientName) <> 0) Then
        
        m_ListFontTitle.AttachToDC bufferDC
        m_ListFontTitle.SetFontColor txtTitleColor
        m_ListFontTitle.FastRenderText offsetX + 0, offsetY + 0, m_GradientCollection(itemIndex).gcGradient.GetGradientName()
        m_ListFontTitle.ReleaseFromDC
        
        'Some special gradients (e.g. "Foreground to transparent") do not have paths; render nothing for these
        If (m_GradientCollection(itemIndex).gcIsSpecial = gs_None) Then
            m_ListFont.AttachToDC bufferDC
            m_ListFont.SetFontColor txtDescriptionColor
            offsetY = offsetY + Interface.FixDPI(5) + m_ListFont.GetHeightOfString(m_GradientCollection(itemIndex).gcGradient.GetGradientName())
            m_ListFont.FastRenderTextWithClipping offsetX, offsetY, tmpRectF.Width - offsetX, tmpRectF.Height, m_GradientCollection(itemIndex).gcPath, True, False, False
            m_ListFont.ReleaseFromDC
        End If
        
    End If
    
End Sub

'Load a gradient collection preview for the first time.  This function is time-consuming; do not call it more
' than absolutely necessary.
Private Sub LoadGradientCollectionPreview(ByVal itemIndex As Long)
    
    With m_GradientCollection(itemIndex)
        If .gcLoadAttempted Then Exit Sub
        Set .gcGradient = New pd2DGradient
        .gcGradientLoadedOK = .gcGradient.LoadGradientFromFile(UserPrefs.GetGradientPath(True) & .gcPath)
        .gcLoadAttempted = True
    End With
    
End Sub

'Load a gradient collection preview for the first time.  This function is time-consuming; do not call it more
' than absolutely necessary.
Private Sub LoadGradientCollectionPreviewDIB(ByVal itemIndex As Long)
    
    'Generate a preview DIB for this gradient.
    If (m_GradientCollection(itemIndex).gcThumb Is Nothing) Then
        
        With m_GradientCollection(itemIndex)
            
            Dim thumbHeight As Long
            thumbHeight = Int(Interface.FixDPIFloat(BLOCKHEIGHT) - 8)
            Set .gcThumb = New pdDIB
            .gcThumb.CreateBlank Interface.FixDPI(GC_THUMB_WIDTH), Interface.FixDPI(thumbHeight), 32, 0, 0
            GDI_Plus.GDIPlusFillDIBRect_Pattern .gcThumb, 0, 0, .gcThumb.GetDIBWidth, .gcThumb.GetDIBHeight, g_CheckerboardPattern
            
            If .gcGradientLoadedOK And (.gcGradient.GetNumOfNodes > 1) Then
                
                'Create a gradient brush and boundary rect
                Dim cBrush As pd2DBrush
                Set cBrush = New pd2DBrush
                cBrush.SetBrushMode P2_BM_Gradient
                cBrush.SetBrushGradientAllSettings .gcGradient.GetGradientAsString()
                cBrush.SetBrushGradientAngle 30!
                cBrush.SetBrushGradientShape P2_GS_Linear
                
                Dim gRectF As RectF
                gRectF.Left = 0!
                gRectF.Top = 0!
                gRectF.Width = .gcThumb.GetDIBWidth
                gRectF.Height = .gcThumb.GetDIBHeight
                cBrush.SetBoundaryRect gRectF
                
                'Render the gradient preview
                Dim cSurface As pd2DSurface
                Drawing2D.QuickCreateSurfaceFromDIB cSurface, .gcThumb, False
                PD2D.FillRectangleF_FromRectF cSurface, cBrush, gRectF
                
            End If
            
            .gcThumb.SetAlphaPremultiplication True
            
        End With
        
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
    If minDistance < (GRADIENT_NODE_WIDTH / 2) / CDbl(picNodePreview.GetWidth) Then
        GetPointAtPosition = minIndex
    Else
        GetPointAtPosition = -1
    End If
    
End Function

'Given an (x, y) position on the gradient interaction window, convert it to the [0, 1] range used by the gradient control.
Private Function ConvertPixelCoordsToNodeCoords(ByVal x As Long) As Single
    
    'Start by converting the current x-position into the range [0, 1]
    Dim uiMin As Single, uiMax As Single, uiRange As Single
    
    'Because the interactive node box is slightly wider than the preview box above it (so that we can center
    ' the top "triangle point" of nodes at the edges of the box), we need to manually solve for the difference
    ' in position between the two picture boxes.  To make it work on high-DPI devices, we bypass VB6's internal
    ' measurement systems and drop into WAPI instead.
    Dim nodePreviewRect As winRect
    If (Not g_WindowManager Is Nothing) Then
        
        g_WindowManager.GetWindowRect_API picNodePreview.hWnd, nodePreviewRect
        
        Dim tmpPoint As PointAPI
        tmpPoint.x = nodePreviewRect.x1
        tmpPoint.y = nodePreviewRect.y1
        
        g_WindowManager.GetScreenToClient Me.hWnd, tmpPoint
        
        uiMin = tmpPoint.x + 1
        uiMax = tmpPoint.x + (nodePreviewRect.x2 - nodePreviewRect.x1) - 2
    
    'This branch should never trigger (as g_WindowManager will always exist), but I've left it here as a
    ' (non-DPI friendly) failsafe.
    Else
        uiMin = picNodePreview.GetLeft + 1
        uiMax = picNodePreview.GetLeft + picNodePreview.GetWidth - 2
    End If
    
    uiRange = uiMax - uiMin
    ConvertPixelCoordsToNodeCoords = (CSng(x) - uiMin) / uiRange
    
    If (ConvertPixelCoordsToNodeCoords < 0!) Then
        ConvertPixelCoordsToNodeCoords = 0
    ElseIf (ConvertPixelCoordsToNodeCoords > 1!) Then
        ConvertPixelCoordsToNodeCoords = 1!
    End If
    
End Function

'When a new active node is selected (or its parameters somehow changed), call this sub to synchronize all UI elements to that node's properties.
Private Sub SyncUIToActiveNode()
    
    If PDMain.IsProgramRunning() Then
    
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

    If PDMain.IsProgramRunning() Then
        
        'Each node is basically comprised of three parts:
        ' 1) An upward arrowhead pointing at the gradient's precise position
        ' 2) a colored block representing the gradient's pure color.  (Opacity is ignored for this UI element)
        ' 3) An outline encompassing (1) and (2), which is colored based on the node's hover state
        
        'To simplify things, we assemble generic paths for (1) and (2), then simply translate and draw them for each individual node.
        Dim baseArrow As pd2DPath, baseBlock As pd2DPath
        Set baseArrow = New pd2DPath
        Set baseBlock = New pd2DPath
        
        'The base arrow is centered at 0, for convenience when translating
        Dim triangleHalfWidth As Single, triangleHeight As Single
        triangleHalfWidth = (GRADIENT_NODE_WIDTH / 2)
        triangleHeight = (picInteract.GetHeight - GRADIENT_NODE_HEIGHT) - 1
        baseArrow.AddTriangle -1 * triangleHalfWidth, triangleHeight, 0, 0, triangleHalfWidth, triangleHeight
        
        'Next up is the colored block, also centered horizontally around position 0
        baseBlock.AddRectangle_Relative -1 * GRADIENT_NODE_WIDTH \ 2, triangleHeight, GRADIENT_NODE_WIDTH, GRADIENT_NODE_HEIGHT
        
        'We also want some duplicate nodes, to remove the need to reset our base node shapes between draws
        Dim tmpArrow As pd2DPath, tmpBlock As pd2DPath
        Set tmpArrow = New pd2DPath
        Set tmpBlock = New pd2DPath
        
        'Finally, some generic scale factors to simplify the process of positioning nodes (who store their positions on the range [0, 1])
        Dim hOffset As Single, hScaleFactor As Single
        hOffset = (picNodePreview.GetLeft - picInteract.GetLeft) + 1
        hScaleFactor = (picNodePreview.GetWidth - 1)
        
        '...and pen/fill objects for the actual rendering
        Dim blockFill As pd2DBrush
        Set blockFill = New pd2DBrush
        blockFill.SetBrushMode P2_BM_Solid
        blockFill.SetBrushOpacity 100#
        
        'Prep the target interaction DIB
        If (m_InteractiveDIB Is Nothing) Then Set m_InteractiveDIB = New pdDIB
        If (m_InteractiveDIB.GetDIBWidth <> Me.picInteract.GetWidth) Or (m_InteractiveDIB.GetDIBHeight <> Me.picInteract.GetHeight) Then
            m_InteractiveDIB.CreateBlank Me.picInteract.GetWidth, Me.picInteract.GetHeight, 24, 0
        Else
            m_InteractiveDIB.ResetDIB
        End If
        
        Dim cSurface As pd2DSurface, cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, m_InteractiveDIB.GetDIBDC, False
        
        'Fill the interaction DIB with the current background color
        Drawing2D.QuickCreateSolidBrush cBrush, g_Themer.GetGenericUIColor(UI_Background)
        PD2D.FillRectangleF cSurface, cBrush, 0, 0, m_InteractiveDIB.GetDIBWidth, m_InteractiveDIB.GetDIBHeight
        cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
        
        'To help the user understand where the interactive area lies, paint a light textured background
        Dim edgePadding As Single
        edgePadding = Interface.FixDPIFloat(10)
        
        Dim patBrush As pd2DBrush
        Set patBrush = New pd2DBrush
        patBrush.SetBrushMode P2_BM_Pattern
        patBrush.SetBrushPattern1Color g_Themer.GetGenericUIColor(UI_GrayLight)
        patBrush.SetBrushPattern1Opacity 50!
        patBrush.SetBrushPattern2Color g_Themer.GetGenericUIColor(UI_Background)
        patBrush.SetBrushPattern2Opacity 100!
        patBrush.SetBrushPatternStyle P2_PS_Weave
        PD2D.FillRectangleF cSurface, patBrush, edgePadding, 0!, m_InteractiveDIB.GetDIBWidth - edgePadding * 2!, m_InteractiveDIB.GetDIBHeight
        
        'Similarly, outline the interactive area
        Dim cOutlinePen As pd2DPen
        Drawing2D.QuickCreateSolidPen cOutlinePen, 1, g_Themer.GetGenericUIColor(UI_GrayLight)
        PD2D.DrawRectangleF cSurface, cOutlinePen, edgePadding, 0!, (m_InteractiveDIB.GetDIBWidth - 1) - edgePadding * 2!, (m_InteractiveDIB.GetDIBHeight - 1)
        
        'If a gradient node is *not* currently hovered, but the mouse lies over the interactive area,
        ' paint a circle to help the user know that interesting stuff happens here.
        If (m_CurHoverPoint < 0) And (m_CurHoverX >= 0) Then
            
            'Limit the hover positioning to the artifical "edges" of the interactive area
            Dim hovIconRadius As Single
            hovIconRadius = Interface.FixDPIFloat(5)
            
            Dim hovLeftBound As Single, hovRightBound As Single
            hovLeftBound = edgePadding + hovIconRadius + Interface.FixDPIFloat(2)
            hovRightBound = (m_InteractiveDIB.GetDIBWidth - 1) - edgePadding - hovIconRadius - Interface.FixDPIFloat(2)
            
            Dim hovRenderX As Single
            If (m_CurHoverX < hovLeftBound) Then
                hovRenderX = hovLeftBound
            ElseIf (m_CurHoverX > hovRightBound) Then
                hovRenderX = hovRightBound
            Else
                hovRenderX = m_CurHoverX
            End If
            
            Dim cBrushAccent As pd2DBrush
            Drawing2D.QuickCreateSolidBrush cBrushAccent, g_Themer.GetGenericUIColor(UI_Accent)
            PD2D.FillCircleF cSurface, cBrushAccent, hovRenderX, m_InteractiveDIB.GetDIBHeight \ 2, hovIconRadius
        
        End If
        
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
            blockFill.SetBrushColor m_GradientPoints(i).PointRGB
            PD2D.FillPath cSurface, blockFill, tmpBlock
            
            'All other renders vary by hover state
            If ((i = m_CurPoint) Or (i = m_CurHoverPoint)) Then
                PD2D.DrawPath cSurface, m_activeOutlinePen, tmpBlock
                PD2D.FillPath cSurface, m_activeArrowFill, tmpArrow
                PD2D.DrawPath cSurface, m_activeOutlinePen, tmpArrow
            Else
                PD2D.DrawPath cSurface, m_inactiveOutlinePen, tmpBlock
                PD2D.FillPath cSurface, m_inactiveArrowFill, tmpArrow
                PD2D.DrawPath cSurface, m_inactiveOutlinePen, tmpArrow
            End If
            
        Next i
        
        Set cSurface = Nothing: Set cBrush = Nothing
        
        'Finally, flip the DIB to the screen
        picInteract.RequestRedraw True
        
    End If

End Sub

'Some user interactions require us to redraw just about everything on the dialog.  Use this shortcut function to do so.
Private Sub RedrawEverything()
    UpdateGradientObjects
    DrawGradientNodes
    UpdatePreview
End Sub

Private Sub picAutoPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_AutoPreviewDIB Is Nothing) Then m_AutoPreviewDIB.AlphaBlendToDC targetDC
End Sub

Private Sub picInteract_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_InteractiveDIB Is Nothing) Then m_InteractiveDIB.AlphaBlendToDC targetDC
End Sub

Private Sub picInteract_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

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
            If (m_CurPoint >= m_NumOfGradientPoints) Then
                m_CurPoint = -1
                SyncUIToActiveNode
            End If
            
        End If
        
    'If this is not an existing point, create a new one now.
    Else
    
        'Make sure the *left* mouse button was clicked
        If ((Button And pdLeftButton) <> 0) Then
            
            'Enlarge the target array as necessary
            If (m_NumOfGradientPoints >= UBound(m_GradientPoints)) Then ReDim Preserve m_GradientPoints(0 To m_NumOfGradientPoints * 2) As GradientPoint
            
            With m_GradientPoints(m_NumOfGradientPoints)
                
                .PointOpacity = 100
                .PointPosition = ConvertPixelCoordsToNodeCoords(x)
                
                'Preset the RGB value to match whatever the gradient already is at this point
                Dim newRGBA As RGBQuad
                m_NodePreview.GetColorAtPosition_RGBA .PointPosition, newRGBA
                .PointRGB = RGB(newRGBA.Red, newRGBA.Green, newRGBA.Blue)
                
            End With
            
            m_CurPoint = m_NumOfGradientPoints
            m_NumOfGradientPoints = m_NumOfGradientPoints + 1
        
        End If
        
    End If
    
    'Regardless of outcome, we need to resync the UI to the active node, and redraw the interaction area and preview
    SyncUIToActiveNode
    UpdatePreview
    DrawGradientNodes

End Sub

Private Sub picInteract_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    picInteract.SetCursorCustom IDC_HAND
End Sub

Private Sub picInteract_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_CurHoverPoint = -1
    m_CurHoverX = -1
    picInteract.SetCursorCustom IDC_DEFAULT
    DrawGradientNodes
End Sub

Private Sub picInteract_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'First, separate our handling by mouse button state
    If (Button And pdLeftButton) <> 0 Then
        
        m_CurHoverX = -1
        
        'The left mouse button is down.  Assign the new position to the active node.
        If (m_CurPoint >= 0) Then
            If chkDistributeEvenly.Value Then chkDistributeEvenly.Value = False
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
        If (tmpPoint <> m_CurHoverPoint) Then
            m_CurHoverPoint = tmpPoint
            m_CurHoverX = -1
        Else
            m_CurHoverX = x
        End If
        
        DrawGradientNodes
    
    End If

End Sub

Private Sub picNodePreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_NodePreviewDIB Is Nothing) Then m_NodePreviewDIB.AlphaBlendToDC targetDC
End Sub

Private Sub sldDensityAuto_Change()
    UpdatePreview
End Sub

Private Sub sldOpacityAuto_Change(Index As Integer)
    UpdatePreview
End Sub

Private Sub sldVaryHSV_Change(Index As Integer)
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
        m_GradientPoints(m_CurPoint).PointPosition = sltNodePosition.Value * 0.01
        RedrawEverything
    End If
End Sub

'The two different gradient panels (manual and auto) require different preview code, obviously.  Call this function
' and it will automatically request a redraw from the active panel.
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then
        If (btsEdit.ListIndex = 1) Then
            UpdatePreview_Manual
        ElseIf (btsEdit.ListIndex = 2) Then
            UpdatePreview_Auto
        End If
    End If
End Sub

Private Sub UpdatePreview_Auto()
    
    If m_SuspendUI Then Exit Sub
    
    'From our automatic gradient settings (currently a start color, end color, and "roughness" setting),
    ' we want to create an *actual* set of gradient points.  This set will be used as the final value for
    ' this gradient, so we need to make sure it is reproducible.  (Hence the use of pdRandomize.)
    If (m_Random Is Nothing) Then
        Set m_Random = New pdRandomize
        m_Random.SetSeed_AutomaticAndRandom
        m_RandomKey = m_Random.GetSeed()
    Else
        m_Random.SetSeed_Float m_RandomKey
    End If
    
    Dim initColor As Long, initOpacity As Single, finalColor As Long, finalOpacity As Single
    initColor = csColorAuto(0).Color
    finalColor = csColorAuto(1).Color
    initOpacity = sldOpacityAuto(0).Value
    finalOpacity = sldOpacityAuto(1).Value
    
    'Our points eventually need to be assembled into a fixed array.  The size of this array varies depending on
    ' the "density" value specified by the user.
    Dim gradPoints() As GradientPoint
    Dim numGradientPoints As Long
    
    'The "density" (e.g. number of auto-generated noise nodes) is displayed to the user on a [0, 100] scale.
    ' We can remap this to effectively any value we want, but there are diminishing returns as the number
    ' climbs.  Currently we limit this to 200 nodes total, which is still probably overkill as a 1000px
    ' gradient only leaves an average of 5 pixels per node, but we can always bump this number up in the
    ' future if users request it.
    numGradientPoints = sldDensityAuto.Value * 2#
    If (numGradientPoints < 1) Then numGradientPoints = 1
    
    'Insert two extra points for our initial and final colors
    ReDim gradPoints(0 To numGradientPoints + 1) As GradientPoint
    
    'Because the gradient class sorts colors automatically, we don't need to handle it here.  As such, we can add
    ' our initial and final colors first.
    Dim i As Long
    For i = 0 To 1
        With gradPoints(i)
            .PointRGB = csColorAuto(i).Color
            .PointOpacity = sldOpacityAuto(i).Value
            .PointPosition = i
        End With
    Next i
    
    Dim r1 As Single, r2 As Single, g1 As Single, g2 As Single, b1 As Single, b2 As Single, a1 As Single, a2 As Single
    a1 = gradPoints(0).PointOpacity
    a2 = gradPoints(1).PointOpacity
    
    r1 = Colors.ExtractRed(gradPoints(0).PointRGB)
    g1 = Colors.ExtractGreen(gradPoints(0).PointRGB)
    b1 = Colors.ExtractBlue(gradPoints(0).PointRGB)
    
    r2 = Colors.ExtractRed(gradPoints(1).PointRGB)
    g2 = Colors.ExtractGreen(gradPoints(1).PointRGB)
    b2 = Colors.ExtractBlue(gradPoints(1).PointRGB)
    
    'Change the variance values from [0, 100] to [0.0, 1.0] scale.
    Dim varyHue As Single, varySaturation As Single, varyValue As Single, varyAlpha As Single
    varyHue = (sldVaryHSV(0).Value * 0.005)
    varySaturation = (sldVaryHSV(1).Value * 0.005)
    varyValue = (sldVaryHSV(2).Value * 0.005)
    varyAlpha = (sldVaryHSV(3).Value * 0.005)
    
    'Next, we want to populate all intermediate points.  We generate these by first calculating what a given point's
    ' values would be if it were non-random; then we vary the point's HSV values according to the user's settings.
    Const ONE_DIV_255 As Double = 1# / 255#
    
    Dim ptIndex As Double, ptRed As Double, ptGreen As Double, ptBlue As Double, ptAlpha As Double
    Dim ptHue As Double, ptSaturation As Double, ptValue As Double
    For i = 2 To UBound(gradPoints)
        
        'Retrieve a random value on the range [0.0, 1.0]
        Do
            ptIndex = m_Random.GetRandomFloat_WH()
        Loop While (ptIndex = 0#) Or (ptIndex = 1#)
        
        'Calculate this point's theoretically "perfect" color and opacity values
        ptRed = (r2 * ptIndex) + (r1 * (1# - ptIndex))
        ptGreen = (g2 * ptIndex) + (g1 * (1# - ptIndex))
        ptBlue = (b2 * ptIndex) + (b1 * (1# - ptIndex))
        ptAlpha = (a2 * ptIndex) + (a1 * (1# - ptIndex))
        
        'Convert the RGB values to HSV (but leave alpha as it is!)
        Colors.fRGBtoHSV ptRed * ONE_DIV_255, ptGreen * ONE_DIV_255, ptBlue * ONE_DIV_255, ptHue, ptSaturation, ptValue
        
        'Vary each point by its "variance" value (including alpha)
        ptHue = ptHue + (varyHue * (1# - m_Random.GetRandomFloat_WH() * 2#))
        ptSaturation = ptSaturation + (varySaturation * (1# - m_Random.GetRandomFloat_WH() * 2#))
        ptValue = ptValue + (varyValue * (1# - m_Random.GetRandomFloat_WH() * 2#))
        ptAlpha = ptAlpha + (varyAlpha * (100# - m_Random.GetRandomFloat_WH() * 200#))
        
        'Clamp accordingly
        If (ptHue < 0#) Then ptHue = ptHue + 1# Else If (ptHue > 1#) Then ptHue = (ptHue - 1#)
        If (ptSaturation < 0#) Then ptSaturation = -1# * ptSaturation Else If (ptSaturation > 1#) Then ptSaturation = 1# - (ptSaturation - 1#)
        If (ptValue < 0#) Then ptValue = -1# * ptValue Else If (ptValue > 1#) Then ptValue = 1# - (ptValue - 1#)
        If (ptAlpha < 0#) Then ptAlpha = -1# * ptAlpha Else If (ptAlpha > 100#) Then ptAlpha = 100# - (ptAlpha - 100#)
        
        'Convert back to RGB
        Colors.fHSVtoRGB ptHue, ptSaturation, ptValue, ptRed, ptGreen, ptBlue
        
        With gradPoints(i)
        
            .PointPosition = ptIndex
            .PointRGB = RGB(Int(ptRed * 255#), Int(ptGreen * 255#), Int(ptBlue * 255#))
            
            'Remember that the user doesn't have to vary alpha; it's available as a toggle
            .PointOpacity = ptAlpha
            
        End With
        
    Next i
    
    'We now have a completed gradient array.  Construct a gradient object and pass it the array.
    If (m_AutoPreview Is Nothing) Then Set m_AutoPreview = New pd2DGradient
    m_AutoPreview.SetGradientShape P2_GS_Linear
    m_AutoPreview.CreateGradientFromPointCollection numGradientPoints + 2, gradPoints
    
    'Render the finished gradient to the sample picture box.
    Dim boundsRect As RectF
    With boundsRect
        .Left = 0
        .Top = 0
        .Width = picAutoPreview.GetWidth
        .Height = picAutoPreview.GetHeight
    End With
    
    If (m_AutoPreviewDIB Is Nothing) Then Set m_AutoPreviewDIB = New pdDIB
    m_AutoPreviewDIB.CreateBlank Me.picAutoPreview.GetWidth, Me.picAutoPreview.GetHeight, 24, 0
    
    Dim cSurface As pd2DSurface, cBrush As pd2DBrush
    Drawing2D.QuickCreateSurfaceFromDIB cSurface, m_AutoPreviewDIB, False
    
    Set cBrush = New pd2DBrush
    cBrush.SetBrushMode P2_BM_Gradient
    cBrush.SetBrushGradientAllSettings m_AutoPreview.GetGradientAsString
    cBrush.SetBoundaryRect boundsRect
    
    With m_AutoPreviewDIB
        PD2D.FillRectangleF cSurface, g_CheckerboardBrush, 0, 0, .GetDIBWidth, .GetDIBHeight
        PD2D.FillRectangleF cSurface, cBrush, 0, 0, .GetDIBWidth, .GetDIBHeight
    End With
    
    'Finish by drawing a neutral border around the outside
    Dim cBorderPen As pd2DPen
    Set cBorderPen = New pd2DPen
    If (Not g_Themer Is Nothing) Then cBorderPen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
    cBorderPen.SetPenWidth 1!
    cBorderPen.SetPenLineJoin P2_LJ_Miter
    PD2D.DrawRectangleI cSurface, cBorderPen, 0, 0, m_AutoPreviewDIB.GetDIBWidth - 1, m_AutoPreviewDIB.GetDIBHeight - 1
    
    Set cSurface = Nothing
    picAutoPreview.RequestRedraw True
            
    'Notify our parent of the update
    If Not (parentGradientControl Is Nothing) Then parentGradientControl.NotifyOfLiveGradientChange GetGradientAsOriginalShape()
    
End Sub

Private Sub UpdatePreview_Manual()

    If (Not m_SuspendUI) Then
    
        If (m_NodePreview.GetNumOfNodes > 0) Then
        
            'Make sure our gradient objects are up-to-date
            UpdateGradientObjects
            
            'Next, use the current gradient nodes to paint a matching preview across the node editor window
            Dim boundsRect As RectF
            
            With boundsRect
                .Left = 0
                .Top = 0
                .Width = picNodePreview.GetWidth
                .Height = picNodePreview.GetHeight
            End With
            
            If (m_NodePreviewDIB Is Nothing) Then Set m_NodePreviewDIB = New pdDIB
            If (m_NodePreviewDIB.GetDIBWidth <> Me.picNodePreview.GetWidth) Or (m_NodePreviewDIB.GetDIBHeight <> Me.picNodePreview.GetHeight) Then
                m_NodePreviewDIB.CreateBlank Me.picNodePreview.GetWidth, Me.picNodePreview.GetHeight, 24, 0
            Else
                m_NodePreviewDIB.ResetDIB
            End If
            
            Dim cSurface As pd2DSurface, cBrush As pd2DBrush
            Drawing2D.QuickCreateSurfaceFromDC cSurface, m_NodePreviewDIB.GetDIBDC, False
            
            Set cBrush = New pd2DBrush
            cBrush.SetBrushMode P2_BM_Gradient
            cBrush.SetBrushGradientAllSettings m_NodePreview.GetGradientAsString
            cBrush.SetBoundaryRect boundsRect
            
            With m_NodePreviewDIB
                PD2D.FillRectangleF cSurface, g_CheckerboardBrush, 0, 0, .GetDIBWidth, .GetDIBHeight
                PD2D.FillRectangleF cSurface, cBrush, 0, 0, .GetDIBWidth, .GetDIBHeight
            End With
            
            'Finish by drawing a neutral border around the outside
            Dim cBorderPen As pd2DPen
            Set cBorderPen = New pd2DPen
            If (Not g_Themer Is Nothing) Then cBorderPen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
            cBorderPen.SetPenWidth 1!
            cBorderPen.SetPenLineJoin P2_LJ_Miter
            PD2D.DrawRectangleI cSurface, cBorderPen, 0, 0, m_NodePreviewDIB.GetDIBWidth - 1, m_NodePreviewDIB.GetDIBHeight - 1
            
            Set cSurface = Nothing
            picNodePreview.RequestRedraw True
                    
            'Notify our parent of the update
            If Not (parentGradientControl Is Nothing) Then parentGradientControl.NotifyOfLiveGradientChange GetGradientAsOriginalShape()
        
        End If
        
    End If
    
End Sub

'Before the custom list box does any painting, we need to retrieve relevant colors from PD's primary theming class.
' Note that this step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor cl_TitleSelected, "TitleSelected", IDE_GRAY
        .LoadThemeColor cl_TitleUnselected, "TitleUnselected", IDE_GRAY
        .LoadThemeColor cl_DescriptionSelected, "TitleSelected", IDE_GRAY
        .LoadThemeColor cl_DescriptionUnselected, "TitleUnselected", IDE_GRAY
    End With
End Sub
