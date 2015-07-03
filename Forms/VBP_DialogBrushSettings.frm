VERSION 5.00
Begin VB.Form dialog_FillSettings 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Fill settings"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12270
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
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   818
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBrushPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   799
      TabIndex        =   2
      Top             =   480
      Width           =   12015
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   7155
      Width           =   12270
      _extentx        =   21643
      _extenty        =   1323
      font            =   "VBP_DialogBrushSettings.frx":0000
      backcolor       =   14802140
      autoloadlastpreset=   -1  'True
      dontautounloadparent=   -1  'True
      dontresetautomatically=   -1  'True
   End
   Begin PhotoDemon.buttonStrip btsStyle 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   12015
      _extentx        =   21193
      _extenty        =   1085
      fontsize        =   12
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   0
      Left            =   120
      Top             =   1200
      Width           =   9375
      _extentx        =   16536
      _extenty        =   556
      caption         =   "fill style:"
      fontsize        =   12
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   9375
      _extentx        =   16536
      _extenty        =   556
      caption         =   "preview:"
      fontsize        =   12
   End
   Begin PhotoDemon.sliderTextCombo sltFillOpacity 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   3240
      Width           =   6060
      _extentx        =   4868
      _extenty        =   873
      max             =   100
      value           =   100
      notchposition   =   2
      notchvaluecustom=   100
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   6
      Left            =   6120
      Top             =   2400
      Width           =   6015
      _extentx        =   10610
      _extenty        =   556
      caption         =   "common settings"
      fontsize        =   12
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   7
      Left            =   6120
      Top             =   2880
      Width           =   6015
      _extentx        =   10610
      _extenty        =   556
      caption         =   "fill opacity"
      fontsize        =   12
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   1
      Left            =   120
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   4
      Top             =   2400
      Width           =   5775
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   5655
         _extentx        =   16536
         _extenty        =   556
         caption         =   "pattern fill settings"
         fontsize        =   12
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   285
         Index           =   14
         Left            =   0
         Top             =   525
         Width           =   5565
         _extentx        =   9816
         _extenty        =   503
         caption         =   "pattern:"
         fontsize        =   12
         forecolor       =   0
      End
      Begin PhotoDemon.pdComboBox cboFillPattern 
         Height          =   450
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   5430
         _extentx        =   9578
         _extenty        =   794
         fontsize        =   12
      End
      Begin PhotoDemon.colorSelector csPattern 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   5535
         _extentx        =   9763
         _extenty        =   873
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   285
         Index           =   11
         Left            =   0
         Top             =   1560
         Width           =   5685
         _extentx        =   10028
         _extenty        =   503
         caption         =   "pattern color and opacity:"
         fontsize        =   12
         forecolor       =   0
      End
      Begin PhotoDemon.colorSelector csPattern 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   5535
         _extentx        =   9763
         _extenty        =   873
         curcolor        =   0
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   285
         Index           =   13
         Left            =   0
         Top             =   3120
         Width           =   5685
         _extentx        =   10028
         _extenty        =   503
         caption         =   "background color and opacity:"
         fontsize        =   12
         forecolor       =   0
      End
      Begin PhotoDemon.sliderTextCombo sltPatternOpacity 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   2520
         Width           =   5700
         _extentx        =   10054
         _extenty        =   873
         max             =   100
         value           =   100
         notchposition   =   2
         notchvaluecustom=   100
      End
      Begin PhotoDemon.sliderTextCombo sltPatternOpacity 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   4080
         Width           =   5700
         _extentx        =   10054
         _extenty        =   873
         max             =   100
         value           =   100
         notchposition   =   2
         notchvaluecustom=   100
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   0
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   3
      Top             =   2400
      Width           =   5775
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   9375
         _extentx        =   16536
         _extenty        =   556
         caption         =   "solid fill settings"
         fontsize        =   12
      End
      Begin PhotoDemon.colorSelector csFillColor 
         Height          =   840
         Left            =   0
         TabIndex        =   7
         Top             =   960
         Width           =   5655
         _extentx        =   9975
         _extenty        =   1482
         curcolor        =   0
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   285
         Index           =   6
         Left            =   0
         Top             =   540
         Width           =   5685
         _extentx        =   10028
         _extenty        =   503
         caption         =   "color:"
         fontsize        =   12
         forecolor       =   0
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   3
      Left            =   120
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   6
      Top             =   2400
      Width           =   5775
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   9375
         _extentx        =   16536
         _extenty        =   556
         caption         =   "texture fill settings"
         fontsize        =   12
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   2
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   5
      Top             =   2400
      Width           =   5775
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   9375
         _extentx        =   16536
         _extenty        =   556
         caption         =   "gradient fill settings"
         fontsize        =   12
      End
   End
End
Attribute VB_Name = "dialog_FillSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Brush Selection Dialog
'Copyright 2015-2015 by Tanner Helland
'Created: 30/June/15 (but assembled from many bits written earlier)
'Last updated: 30/June/15
'Last update: start migrating brush creation bits into this singular dialog
'
'Comprehensive brush selection dialog.  This dialog is currently based around the properties of GDI+ brushes, but it could
' easily be expanded in the future due to its modular design.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'OK/Cancel result from the dialog
Private userAnswer As VbMsgBoxResult

'The original brush when the dialog was first loaded
Private m_OldBrush As String

'Brush strings are generated with the help of a fill (GDI+ brush) class.  This class also renders a preview of the current fill.
Private m_Filler As pdGraphicsBrush

'If a user control spawned this dialog, it will pass itself as a reference.  We can then send brush updates back
' to the control, allowing for real-time updates on the screen despite a modal dialog being raised!
Private parentBrushControl As brushSelector

'Recently used brushes are loaded to/saved from a custom XML file
Private m_XMLEngine As pdXML

'The file where we'll store recent brush data when the program is closed.  (At present, this file is located in PD's
' /Data/Presets/ folder.
Private m_XMLFilename As String

'Brush preview DIB
Private m_PreviewDIB As pdDIB

'To prevent recursive setting changes, this value can be set to TRUE to prevent live preview updates
Private m_SuspendRedraws As Boolean

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The newly selected brush (if any) is returned via this property
Public Property Get newBrush() As String
    newBrush = m_Filler.getBrushAsString
End Property

'The ShowDialog routine presents the user with this form.
Public Sub showDialog(ByVal initialBrush As String, Optional ByRef callingControl As brushSelector = Nothing)

    Debug.Print "Initial brush=" & initialBrush

    'Store a reference to the calling control (if any)
    Set parentBrushControl = callingControl

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Cache the initial brush parameter so we can access it elsewhere
    m_OldBrush = initialBrush
    Set m_Filler = New pdGraphicsBrush
    m_Filler.createBrushFromString initialBrush
    
    If Len(initialBrush) = 0 Then initialBrush = m_Filler.getBrushAsString
    
    'Sync all controls to the initial brush parameters
    syncControlsToFillObject
    updatePreview
    
    Debug.Print "preview time opacity:" & m_Filler.getBrushProperty(pgbs_PrimaryOpacity) & ", " & sltFillOpacity.Value
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    
    'Apply extra images and tooltips to certain controls
    
    'Apply visual themes
    makeFormPretty Me
    
    'Initialize an XML engine, which we will use to read/write recent brush data to file
    Set m_XMLEngine = New pdXML
    
    'The XML file will be stored in the Preset path (/Data/Presets)
    m_XMLFilename = g_UserPreferences.getPresetPath & "Brush_Selector.xml"
    
    'TODO: if an XML file exists, load its contents now
    'loadRecentBrushList
        
    'Display the dialog
    showPDDialog vbModal, Me, True

End Sub

Private Sub btsStyle_Click(ByVal buttonIndex As Long)
    
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        picContainer(i).Visible = CBool(i = buttonIndex)
    Next i
    
    updatePreview
    
End Sub

Private Sub cboFillPattern_Click()
    updatePreview
End Sub

'CANCEL BUTTON
Private Sub cmdBar_CancelClick()
    userAnswer = vbCancel
    Me.Hide
End Sub

'OK BUTTON
Private Sub cmdBar_OKClick()

    'Store the newBrush value (which the dialog handler will use to return the selected brush to the caller)
    updateFillObject
    
    'TODO: save the current list of recently used brushes
    'saveRecentBrushList
    
    userAnswer = vbOK
    Me.Hide

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Reset our generic fill object
    Set m_Filler = New pdGraphicsBrush
    m_Filler.createBrushFromString ""
    
    'Synchronize all controls to the updated settings
    syncControlsToFillObject
    updatePreview
    
End Sub

Private Sub csFillColor_ColorChanged()
    updatePreview
End Sub

Private Sub csPattern_ColorChanged(Index As Integer)
    updatePreview
End Sub

Private Sub Form_Load()
    
    m_SuspendRedraws = True
    
    'Populate the main style button strip
    btsStyle.AddItem "solid", 0
    btsStyle.AddItem "pattern", 1
    btsStyle.AddItem "gradient", 2
    btsStyle.AddItem "texture", 3
    btsStyle.ListIndex = 0
    btsStyle_Click 0
    
    'Hatch patterns
    
    'TODO: custom pattern dropdown, since we'll be using it elsewhere!
    cboFillPattern.Clear
    cboFillPattern.AddItem "horizontal"
    cboFillPattern.AddItem "vertical"
    cboFillPattern.AddItem "forward diagonal"
    cboFillPattern.AddItem "backward diagonal"
    cboFillPattern.AddItem "cross"
    cboFillPattern.AddItem "diagonal cross"
    cboFillPattern.ListIndex = 0
    
    If g_IsProgramRunning Then
    
        If m_Filler Is Nothing Then Set m_Filler = New pdGraphicsBrush
        Set m_PreviewDIB = New pdDIB
                
    End If
    
    m_SuspendRedraws = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Update our internal brush class against any/all changed settings.
Private Sub updateFillObject()

    With m_Filler
        .setBrushProperty pgbs_BrushMode, btsStyle.ListIndex
        .setBrushProperty pgbs_PrimaryColor, csFillColor.Color
        .setBrushProperty pgbs_PrimaryOpacity, sltFillOpacity.Value
        .setBrushProperty pgbs_PatternID, cboFillPattern.ListIndex
        .setBrushProperty pgbs_PatternColor1, csPattern(0).Color
        .setBrushProperty pgbs_PatternColor1Opacity, sltPatternOpacity(0).Value
        .setBrushProperty pgbs_PatternColor2, csPattern(1).Color
        .setBrushProperty pgbs_PatternColor2Opacity, sltPatternOpacity(1).Value
    End With

End Sub

Private Sub updatePreview()
    
    If Not m_SuspendRedraws Then
    
        'Make sure our fill object is up-to-date
        updateFillObject
        
        'Retrieve a matching brush handle
        Dim gdipBrush As Long
        gdipBrush = m_Filler.getBrushHandle()
        
        'Prep the preview DIB
        If m_PreviewDIB Is Nothing Then Set m_PreviewDIB = New pdDIB
        
        If (m_PreviewDIB.getDIBWidth <> Me.picBrushPreview.ScaleWidth) Or (m_PreviewDIB.getDIBHeight <> Me.picBrushPreview.ScaleHeight) Then
            m_PreviewDIB.createBlank Me.picBrushPreview.ScaleWidth, Me.picBrushPreview.ScaleHeight, 24, 0
        Else
            m_PreviewDIB.resetDIB
        End If
        
        'Create the preview image
        GDI_Plus.GDIPlusFillDIBRect_Pattern m_PreviewDIB, 0, 0, m_PreviewDIB.getDIBWidth, m_PreviewDIB.getDIBHeight, g_CheckerboardPattern
        GDI_Plus.GDIPlusFillDC_Brush m_PreviewDIB.getDIBDC, gdipBrush, 0, 0, Me.picBrushPreview.ScaleWidth, Me.picBrushPreview.ScaleHeight
        
        'Copy the preview image to the screen
        m_PreviewDIB.renderToPictureBox Me.picBrushPreview
        
        'Release our GDI+ handle
        m_Filler.releaseBrushHandle gdipBrush
        
        'Notify our parent of the update
        If Not (parentBrushControl Is Nothing) Then parentBrushControl.notifyOfLiveBrushChange m_Filler.getBrushAsString
        
    End If
    
End Sub

Private Sub sltFillOpacity_Change()
    updatePreview
    Debug.Print "opacity:" & m_Filler.getBrushProperty(pgbs_PrimaryOpacity) & ", " & sltFillOpacity.Value
End Sub

Private Sub sltPatternOpacity_Change(Index As Integer)
    updatePreview
End Sub

Private Sub syncControlsToFillObject()
        
    m_SuspendRedraws = True
        
    With m_Filler
        
        btsStyle.ListIndex = .getBrushProperty(pgbs_BrushMode)
        
        csFillColor.Color = .getBrushProperty(pgbs_PrimaryColor)
        sltFillOpacity.Value = .getBrushProperty(pgbs_PrimaryOpacity)
        
        cboFillPattern.ListIndex = .getBrushProperty(pgbs_PatternID)
        csPattern(0).Color = .getBrushProperty(pgbs_PatternColor1)
        csPattern(1).Color = .getBrushProperty(pgbs_PatternColor2)
        sltPatternOpacity(0).Value = .getBrushProperty(pgbs_PatternColor1Opacity)
        sltPatternOpacity(1).Value = .getBrushProperty(pgbs_PatternColor2Opacity)
    
    End With
    
    m_SuspendRedraws = False
    
End Sub
