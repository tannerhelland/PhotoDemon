VERSION 5.00
Begin VB.Form dialog_FillSettings 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fill settings"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   375
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   818
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBrushPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   300
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   787
      TabIndex        =   2
      Top             =   480
      Width           =   11835
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   7155
      Width           =   12270
      _extentx        =   21643
      _extenty        =   1323
      autoloadlastpreset=   -1  'True
      dontautounloadparent=   -1  'True
      dontresetautomatically=   -1  'True
   End
   Begin PhotoDemon.pdButtonStrip btsStyle 
      Height          =   1050
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   12015
      _extentx        =   21193
      _extenty        =   1852
      caption         =   "fill style"
      fontsize        =   12
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   9375
      _extentx        =   16536
      _extenty        =   556
      caption         =   "preview"
      fontsize        =   12
   End
   Begin PhotoDemon.pdSlider sltFillOpacity 
      CausesValidation=   0   'False
      Height          =   705
      Left            =   6120
      TabIndex        =   8
      Top             =   3000
      Width           =   6060
      _extentx        =   10689
      _extenty        =   1270
      caption         =   "fill opacity"
      max             =   100
      value           =   100
      notchposition   =   2
      notchvaluecustom=   100
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   5
      Left            =   6120
      Top             =   2400
      Width           =   6015
      _extentx        =   10610
      _extenty        =   556
      caption         =   "common settings"
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
      Begin PhotoDemon.pdListBoxOD lstFillPattern 
         Height          =   2535
         Left            =   0
         TabIndex        =   14
         Top             =   480
         Width           =   5655
         _extentx        =   9975
         _extenty        =   4471
         caption         =   "pattern"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   5655
         _extentx        =   16536
         _extenty        =   556
         caption         =   "pattern fill settings"
         fontsize        =   12
      End
      Begin PhotoDemon.pdColorSelector csPattern 
         Height          =   855
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   3120
         Width           =   2775
         _extentx        =   9975
         _extenty        =   1508
         caption         =   "pattern color"
      End
      Begin PhotoDemon.pdColorSelector csPattern 
         Height          =   855
         Index           =   1
         Left            =   2880
         TabIndex        =   10
         Top             =   3120
         Width           =   2775
         _extentx        =   9975
         _extenty        =   1508
         curcolor        =   0
         caption         =   "background color"
      End
      Begin PhotoDemon.pdSlider sltPatternOpacity 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   4080
         Width           =   2820
         _extentx        =   10054
         _extenty        =   873
         max             =   100
         value           =   100
         notchposition   =   2
         notchvaluecustom=   100
      End
      Begin PhotoDemon.pdSlider sltPatternOpacity 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   1
         Left            =   2880
         TabIndex        =   12
         Top             =   4080
         Width           =   2820
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
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   9375
         _extentx        =   16536
         _extenty        =   556
         caption         =   "solid fill settings"
         fontsize        =   12
      End
      Begin PhotoDemon.pdColorSelector csFillColor 
         Height          =   1560
         Left            =   0
         TabIndex        =   7
         Top             =   600
         Width           =   5655
         _extentx        =   9975
         _extenty        =   2752
         curcolor        =   0
         caption         =   "color"
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
         Index           =   4
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
      Begin PhotoDemon.pdGradientSelector gsPrimary 
         Height          =   1335
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   5655
         _extentx        =   9975
         _extenty        =   2355
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   9375
         _extentx        =   16536
         _extenty        =   556
         caption         =   "gradient fill settings"
         fontsize        =   12
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   7
         Left            =   0
         Top             =   600
         Width           =   5565
         _extentx        =   9816
         _extenty        =   503
         caption         =   "gradient"
         fontsize        =   12
         forecolor       =   0
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
'Copyright 2015-2016 by Tanner Helland
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
Private m_Filler As pd2DBrush

'If a user control spawned this dialog, it will pass itself as a reference.  We can then send brush updates back
' to the control, allowing for real-time updates on the screen despite a modal dialog being raised!
Private parentBrushControl As pdBrushSelector

'Recently used brushes are loaded to/saved from a custom XML file
Private m_XMLEngine As pdXML

'The file where we'll store recent brush data when the program is closed.  (At present, this file is located in PD's
' /Data/Presets/ folder.
Private m_XMLFilename As String

'Brush preview DIB
Private m_PreviewDIB As pdDIB

'To prevent recursive setting changes, this value can be set to TRUE to prevent live preview updates
Private m_SuspendRedraws As Boolean

'Hatch count is constant, regardless of OS
Private Const NUM_OF_HATCHES As Long = 53

'2D painting support classes
Private m_Painter As pd2DPainter

'When the form first loads, we find the longest hatch index number (longest in *pixels*).  We do this so that we
' can align all hatch previews identically.
Private m_LargestWidth As Single

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PD_FILL_PATTERN
    [_First] = 0
    PDFP_Background = 0
    PDFP_Caption = 1
    PDFP_ItemBorder = 2
    PDFP_HatchBorder = 3
    [_Last] = 3
    [_Count] = 4
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The newly selected brush (if any) is returned via this property
Public Property Get newBrush() As String
    newBrush = m_Filler.GetBrushPropertiesAsXML
End Property

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByVal initialBrush As String, Optional ByRef callingControl As pdBrushSelector = Nothing)

    Debug.Print "Initial brush=" & initialBrush

    'Store a reference to the calling control (if any)
    Set parentBrushControl = callingControl

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Cache the initial brush parameter so we can access it elsewhere
    m_OldBrush = initialBrush
    Set m_Filler = New pd2DBrush
    m_Filler.SetBrushPropertiesFromXML initialBrush
    If Len(initialBrush) = 0 Then initialBrush = m_Filler.GetBrushPropertiesAsXML
    
    'Sync all controls to the initial brush parameters
    SyncControlsToFillObject
    UpdatePreview
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    
    'Apply extra images and tooltips to certain controls
    
    'Apply visual themes
    ApplyThemeAndTranslations Me
    
    'Initialize an XML engine, which we will use to read/write recent brush data to file
    Set m_XMLEngine = New pdXML
    
    'The XML file will be stored in the Preset path (/Data/Presets)
    m_XMLFilename = g_UserPreferences.GetPresetPath & "Brush_Selector.xml"
    
    'TODO: if an XML file exists, load its contents now
    'loadRecentBrushList
        
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub btsStyle_Click(ByVal buttonIndex As Long)
    
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        picContainer(i).Visible = CBool(i = buttonIndex)
    Next i
    
    UpdatePreview
    
End Sub

'CANCEL BUTTON
Private Sub cmdBar_CancelClick()
    userAnswer = vbCancel
    Me.Hide
End Sub

'OK BUTTON
Private Sub cmdBar_OKClick()

    'Store the newBrush value (which the dialog handler will use to return the selected brush to the caller)
    UpdateFillObject
    
    'TODO: save the current list of recently used brushes
    'saveRecentBrushList
    
    userAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Reset our generic fill object
    Set m_Filler = New pd2DBrush
    m_Filler.SetBrushPropertiesFromXML ""
    
    'Synchronize all controls to the updated settings
    SyncControlsToFillObject
    UpdatePreview
    
End Sub

Private Sub csFillColor_ColorChanged()
    UpdatePreview
End Sub

Private Sub csPattern_ColorChanged(Index As Integer)
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    m_SuspendRedraws = True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PD_FILL_PATTERN: colorCount = [_Count]
    m_Colors.InitializeColorList "PDFillPatterns", colorCount
    UpdateColorList
    
    'Populate the main style button strip
    btsStyle.AddItem "solid", 0
    btsStyle.AddItem "pattern", 1
    btsStyle.AddItem "gradient", 2
    'btsStyle.AddItem "texture", 3      'texture brushes are still TODO!
    btsStyle.ListIndex = 0
    btsStyle_Click 0
    
    'The hatch preview box is owner-drawn, so calculate some additional metrics now
    Drawing2D.QuickCreatePainter m_Painter
    
    Dim tmpFont As pdFont
    Set tmpFont = Font_Management.GetMatchingUIFont(10#)
    
    Dim tmpWidth As Long
    m_LargestWidth = 0
    
    lstFillPattern.ListItemHeight = FixDPI(24)
    lstFillPattern.SetAutomaticRedraws False
    lstFillPattern.Clear
    
    Dim i As Long
    For i = 0 To NUM_OF_HATCHES - 1
        lstFillPattern.AddItem CStr(i), i
        tmpWidth = tmpFont.GetWidthOfString(CStr(i))
        If (tmpWidth > m_LargestWidth) Then m_LargestWidth = tmpWidth
    Next i
    
    lstFillPattern.SetAutomaticRedraws True, True
    
    'Numbers will also have a trailing dash, so add that width now
    m_LargestWidth = m_LargestWidth + tmpFont.GetWidthOfString(" - ")
    
    If g_IsProgramRunning Then
        If (m_Filler Is Nothing) Then Set m_Filler = New pd2DBrush
        Set m_PreviewDIB = New pdDIB
    End If
    
    m_SuspendRedraws = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Update our internal brush class against any/all changed settings.
Private Sub UpdateFillObject()

    With m_Filler
        .SetBrushProperty P2_BrushMode, btsStyle.ListIndex
        .SetBrushProperty P2_BrushColor, csFillColor.Color
        .SetBrushProperty P2_BrushOpacity, sltFillOpacity.Value
        .SetBrushProperty P2_BrushPatternStyle, lstFillPattern.ListIndex
        .SetBrushProperty P2_BrushPattern1Color, csPattern(0).Color
        .SetBrushProperty P2_BrushPattern1Opacity, sltPatternOpacity(0).Value
        .SetBrushProperty P2_BrushPattern2Color, csPattern(1).Color
        .SetBrushProperty P2_BrushPattern2Opacity, sltPatternOpacity(1).Value
        .SetBrushProperty P2_BrushGradientXML, gsPrimary.Gradient
    End With

End Sub

Private Sub UpdatePreview()
    
    If Not m_SuspendRedraws Then
    
        'Make sure our fill object is up-to-date
        UpdateFillObject
        
        'Retrieve a matching brush handle
        Dim gdipBrush As Long, cBounds As RECTF
        
        With cBounds
            .Left = 0
            .Top = 0
            .Width = m_PreviewDIB.GetDIBWidth
            .Height = m_PreviewDIB.GetDIBHeight
        End With
        
        m_Filler.SetBoundaryRect cBounds
        gdipBrush = m_Filler.GetHandle
        
        'Prep the preview DIB
        If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
        If (m_PreviewDIB.GetDIBWidth <> Me.picBrushPreview.ScaleWidth) Or (m_PreviewDIB.GetDIBHeight <> Me.picBrushPreview.ScaleHeight) Then
            m_PreviewDIB.CreateBlank Me.picBrushPreview.ScaleWidth, Me.picBrushPreview.ScaleHeight, 24, 0
        Else
            m_PreviewDIB.ResetDIB
        End If
        
        'Create the preview image
        GDI_Plus.GDIPlusFillDIBRect_Pattern m_PreviewDIB, 0, 0, m_PreviewDIB.GetDIBWidth, m_PreviewDIB.GetDIBHeight, g_CheckerboardPattern
        GDI_Plus.GDIPlusFillDC_Brush m_PreviewDIB.GetDIBDC, gdipBrush, 0, 0, Me.picBrushPreview.ScaleWidth, Me.picBrushPreview.ScaleHeight
        
        'Copy the preview image to the screen
        m_PreviewDIB.RenderToPictureBox Me.picBrushPreview
        
        'Release our GDI+ handle
        m_Filler.ReleaseBrush
        
        'Notify our parent of the update
        If Not (parentBrushControl Is Nothing) Then parentBrushControl.NotifyOfLiveBrushChange m_Filler.GetBrushPropertiesAsXML
        
    End If
    
End Sub

Private Sub gsPrimary_GradientChanged()
    UpdatePreview
End Sub

Private Sub lstFillPattern_Click()
    UpdatePreview
End Sub

Private Sub lstFillPattern_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)
    
    Dim tmpRectF As RECTF
    CopyMemory ByVal VarPtr(tmpRectF), ByVal ptrToRectF, 16&
    
    Dim itemBackColor As Long, itemTextColor As Long, itemBorderColor As Long, hatchBorderColor As Long
    itemBackColor = m_Colors.RetrieveColor(PDFP_Background, Me.Enabled, itemIsSelected, itemIsHovered)
    itemTextColor = m_Colors.RetrieveColor(PDFP_Caption, Me.Enabled, itemIsSelected, itemIsHovered)
    itemBorderColor = m_Colors.RetrieveColor(PDFP_ItemBorder, Me.Enabled, itemIsSelected, itemIsHovered)
    hatchBorderColor = m_Colors.RetrieveColor(PDFP_HatchBorder, Me.Enabled, itemIsSelected, itemIsHovered)
    
    'Fill the background first
    Dim cSurface As pd2DSurface, cBrush As pd2DBrush, cPen As pd2DPen
    Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, False
    Drawing2D.QuickCreateSolidBrush cBrush, itemBackColor
    Drawing2D.QuickCreateSolidPen cPen, 1, itemBorderColor, , P2_LJ_Miter
    
    If (Not (m_Painter Is Nothing)) Then
        
        m_Painter.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
        
        'Next, fill the border
        m_Painter.DrawRectangleF_FromRectF cSurface, cPen, tmpRectF
        
        'Next, draw the caption
        Dim tmpFont As pdFont
        Set tmpFont = Font_Management.GetMatchingUIFont(10#)
        
        Dim tmpString As String
        tmpString = CStr(itemIndex + 1) & " - "
        
        tmpFont.SetFontColor itemTextColor
        tmpFont.AttachToDC bufferDC
        tmpFont.SetTextAlignment vbLeftJustify
        tmpFont.FastRenderTextWithClipping tmpRectF.Left + FixDPI(4), tmpRectF.Top, tmpRectF.Width, tmpRectF.Height, tmpString, False, True, False
        tmpFont.ReleaseFromDC
        Set tmpFont = Nothing
        
        'Finally, draw the hatch
        Dim hatchRect As RECTF
        
        With hatchRect
            .Left = tmpRectF.Left + FixDPI(4) + m_LargestWidth
            .Top = tmpRectF.Top + 2#
            .Height = tmpRectF.Height - 4#
            .Width = (tmpRectF.Left + tmpRectF.Width) - (hatchRect.Left) - FixDPI(4)
        End With
        
        cBrush.ReleaseBrush
        cBrush.SetBrushMode P2_BM_Pattern
        cBrush.SetBrushPatternStyle itemIndex
        cBrush.SetBrushPattern1Color vbBlack
        cBrush.SetBrushPattern1Opacity 100
        cBrush.SetBrushPattern2Color vbWhite
        cBrush.SetBrushPattern2Opacity 100
        cBrush.CreateBrush
        
        cSurface.SetSurfaceRenderingOriginX hatchRect.Left
        cSurface.SetSurfaceRenderingOriginY hatchRect.Top
        m_Painter.FillRectangleF_FromRectF cSurface, cBrush, hatchRect
        
        cPen.SetPenColor hatchBorderColor
        m_Painter.DrawRectangleF_FromRectF cSurface, cPen, hatchRect
    
    End If
    
End Sub

Private Sub sltFillOpacity_Change()
    UpdatePreview
End Sub

Private Sub sltPatternOpacity_Change(Index As Integer)
    UpdatePreview
End Sub

Private Sub SyncControlsToFillObject()
        
    m_SuspendRedraws = True
        
    With m_Filler
        
        btsStyle.ListIndex = .GetBrushProperty(P2_BrushMode)
        
        csFillColor.Color = .GetBrushProperty(P2_BrushColor)
        sltFillOpacity.Value = .GetBrushProperty(P2_BrushOpacity)
        
        lstFillPattern.ListIndex = .GetBrushProperty(P2_BrushPatternStyle)
        csPattern(0).Color = .GetBrushProperty(P2_BrushPattern1Color)
        csPattern(1).Color = .GetBrushProperty(P2_BrushPattern2Color)
        sltPatternOpacity(0).Value = .GetBrushProperty(P2_BrushPattern1Opacity)
        sltPatternOpacity(1).Value = .GetBrushProperty(P2_BrushPattern2Opacity)
        
        gsPrimary.Gradient = .GetBrushProperty(P2_BrushGradientXML)
    
    End With
    
    m_SuspendRedraws = False
    
End Sub

'Before the hatch list box does any painting, we need to retrieve relevant colors from PD's primary theming class.
' Note that this step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDFP_Background, "Background", IDE_WHITE
        .LoadThemeColor PDFP_Caption, "Caption", IDE_GRAY
        .LoadThemeColor PDFP_ItemBorder, "ItemBorder", IDE_WHITE
        .LoadThemeColor PDFP_HatchBorder, "HatchBorder", IDE_GRAY
    End With
End Sub
