VERSION 5.00
Begin VB.Form dialog_FillSettings 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fill settings"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12270
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
   ScaleHeight     =   588
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   818
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdPictureBox picBrushPreview 
      Height          =   1815
      Left            =   270
      Top             =   480
      Width           =   11865
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   8070
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   1323
      AutoloadLastPreset=   -1  'True
      DontAutoUnloadParent=   -1  'True
      DontResetAutomatically=   -1  'True
   End
   Begin PhotoDemon.pdButtonStrip btsStyle 
      Height          =   1050
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1852
      Caption         =   "fill style"
      FontSize        =   12
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   556
      Caption         =   "preview"
      FontSize        =   12
   End
   Begin PhotoDemon.pdSlider sltFillOpacity 
      CausesValidation=   0   'False
      Height          =   705
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   1244
      Caption         =   "opacity"
      Max             =   100
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdContainer ctlGroupFill 
      Height          =   3600
      Index           =   0
      Left            =   120
      Top             =   4440
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6350
      Begin PhotoDemon.pdColorSelector csFillColor 
         Height          =   2535
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4471
         Caption         =   "color"
         curColor        =   0
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupFill 
      Height          =   3600
      Index           =   3
      Left            =   120
      Top             =   4440
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6350
   End
   Begin PhotoDemon.pdContainer ctlGroupFill 
      Height          =   3600
      Index           =   2
      Left            =   120
      Top             =   4440
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6350
      Begin PhotoDemon.pdGradientSelector gsPrimary 
         Height          =   2535
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4471
         Caption         =   "colors"
      End
      Begin PhotoDemon.pdButtonStrip btsGradientShape 
         Height          =   1035
         Left            =   6120
         TabIndex        =   4
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1826
         Caption         =   "shape"
      End
      Begin PhotoDemon.pdSlider sldGradientAngle 
         Height          =   705
         Left            =   6120
         TabIndex        =   5
         Top             =   1080
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1244
         Caption         =   "angle"
         Max             =   360
         SigDigits       =   1
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupFill 
      Height          =   3600
      Index           =   1
      Left            =   120
      Top             =   4440
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6350
      Begin PhotoDemon.pdListBoxOD lstFillPattern 
         Height          =   2535
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4471
         Caption         =   "pattern"
      End
      Begin PhotoDemon.pdColorSelector csPattern 
         Height          =   1095
         Index           =   0
         Left            =   6000
         TabIndex        =   9
         Top             =   0
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   1931
         Caption         =   "pattern color"
      End
      Begin PhotoDemon.pdColorSelector csPattern 
         Height          =   1095
         Index           =   1
         Left            =   9060
         TabIndex        =   10
         Top             =   0
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   1931
         Caption         =   "background color"
         curColor        =   0
      End
      Begin PhotoDemon.pdSlider sltPatternOpacity 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   0
         Left            =   6000
         TabIndex        =   11
         Top             =   1200
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   714
         Max             =   100
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.pdSlider sltPatternOpacity 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   1
         Left            =   9060
         TabIndex        =   2
         Top             =   1200
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   714
         Max             =   100
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
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
'Copyright 2015-2026 by Tanner Helland
'Created: 30/June/15 (but assembled from many bits written earlier)
'Last updated: 13/April/22
'Last update: replace lingering picture box with pdPictureBox
'
'Comprehensive brush selection dialog.  This dialog is currently based around the properties of GDI+ brushes, but it
' could easily be expanded in the future thanks to its modular design.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK/Cancel result from the dialog
Private m_userAnswer As VbMsgBoxResult

'The original brush when the dialog was first loaded
Private m_OldBrush As String

'Brush strings are generated with the help of a fill (GDI+ brush) class.  This class also renders a preview of the current fill.
Private m_Filler As pd2DBrush

'Gradient brushes are constructed with help from a pdGradient instance
Private m_Gradient As pd2DGradient

'If a user control spawned this dialog, it will pass itself as a reference.  We can then send brush updates back
' to the control, allowing for real-time updates on the screen despite a modal dialog being raised!
Private m_parentBrushControl As pdBrushSelector

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
    DialogResult = m_userAnswer
End Property

'The newly selected brush (if any) is returned via this property
Public Property Get newBrush() As String
    newBrush = m_Filler.GetBrushPropertiesAsXML
End Property

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByVal initialBrush As String, Optional ByRef callingControl As pdBrushSelector = Nothing)
    
    'Store a reference to the calling control (if any)
    Set m_parentBrushControl = callingControl

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_userAnswer = vbCancel
    
    'Cache the initial brush parameter so we can access it elsewhere
    m_OldBrush = initialBrush
    Set m_Filler = New pd2DBrush
    m_Filler.SetBrushPropertiesFromXML initialBrush
    If (LenB(initialBrush) = 0) Then initialBrush = m_Filler.GetBrushPropertiesAsXML
    
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
    m_XMLFilename = UserPrefs.GetPresetPath & "Brush_Selector.xml"
    
    'TODO: if an XML file exists, load its contents now
    'LoadRecentBrushList
        
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub btsGradientShape_Click(ByVal buttonIndex As Long)
    UpdateGradientOptionVisibility
    UpdatePreview
End Sub

Private Sub btsStyle_Click(ByVal buttonIndex As Long)
    
    Dim i As Long
    For i = ctlGroupFill.lBound To ctlGroupFill.UBound
        ctlGroupFill(i).Visible = (i = buttonIndex)
    Next i
    
    UpdatePreview
    
End Sub

'CANCEL BUTTON
Private Sub cmdBar_CancelClick()
    m_userAnswer = vbCancel
    Me.Hide
End Sub

'OK BUTTON
Private Sub cmdBar_OKClick()

    'Store the newBrush value (which the dialog handler will use to return the selected brush to the caller)
    UpdateFillObject
    
    'TODO: save the current list of recently used brushes
    'SaveRecentBrushList
    
    m_userAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Reset our generic fill object
    Set m_Filler = New pd2DBrush
    m_Filler.SetBrushPropertiesFromXML vbNullString
    
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
    
    'Populate various button strip selectors
    btsStyle.AddItem "solid", 0
    btsStyle.AddItem "pattern", 1
    btsStyle.AddItem "gradient", 2
    'btsStyle.AddItem "texture", 3      'texture brushes are still TODO!
    btsStyle.ListIndex = 0
    btsStyle_Click 0
    
    btsGradientShape.AddItem "line", 0
    btsGradientShape.AddItem "reflection", 1
    btsGradientShape.AddItem "circle", 2
    btsGradientShape.AddItem "rectangle", 3
    btsGradientShape.AddItem "diamond", 4
    UpdateGradientOptionVisibility
    
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(10#)
    
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
    
    If PDMain.IsProgramRunning() Then
        If (m_Filler Is Nothing) Then Set m_Filler = New pd2DBrush
        If (m_Gradient Is Nothing) Then Set m_Gradient = New pd2DGradient
        Set m_PreviewDIB = New pdDIB
    End If
    
    m_SuspendRedraws = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdateGradientOptionVisibility()
    
    'Show/hide the angle slider depending on the current shape
    If (btsGradientShape.ListIndex = 0) Or (btsGradientShape.ListIndex = 1) Then
        sldGradientAngle.Visible = True
    Else
        sldGradientAngle.Visible = False
    End If
    
End Sub

'Update our internal brush class against any/all changed settings.
Private Sub UpdateFillObject()

    With m_Filler
        .SetBrushMode btsStyle.ListIndex
        .SetBrushColor csFillColor.Color
        .SetBrushOpacity sltFillOpacity.Value
        .SetBrushPatternStyle lstFillPattern.ListIndex
        .SetBrushPattern1Color csPattern(0).Color
        .SetBrushPattern1Opacity sltPatternOpacity(0).Value
        .SetBrushPattern2Color csPattern(1).Color
        .SetBrushPattern2Opacity sltPatternOpacity(1).Value
        
        'Gradient settings are first passed through a pd2DGradient instance, which condenses all the gradient options
        ' into a single settable string.
        m_Gradient.CreateGradientFromString gsPrimary.Gradient
        m_Gradient.SetGradientShape btsGradientShape.ListIndex
        m_Gradient.SetGradientAngle sldGradientAngle.Value
        .SetBrushGradientAllSettings m_Gradient.GetGradientAsString
        
    End With

End Sub

Private Sub UpdatePreview()
    
    If (Not m_SuspendRedraws) Then
    
        'Make sure our fill object is up-to-date
        UpdateFillObject
        
        'Gradient fills require a boundary rect (to establish the "size" of the gradient)
        Dim cBounds As RectF
        
        With cBounds
            .Left = 0
            .Top = 0
            .Width = m_PreviewDIB.GetDIBWidth
            .Height = m_PreviewDIB.GetDIBHeight
        End With
        
        m_Filler.SetBoundaryRect cBounds
        
        'Prep the preview DIB
        If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
        If (m_PreviewDIB.GetDIBWidth <> Me.picBrushPreview.GetWidth) Or (m_PreviewDIB.GetDIBHeight <> Me.picBrushPreview.GetHeight) Then
            m_PreviewDIB.CreateBlank Me.picBrushPreview.GetWidth, Me.picBrushPreview.GetHeight, 24, 0
        Else
            m_PreviewDIB.ResetDIB
        End If
        
        'Render a checkerboard background and brush preview
        Dim cSurface As pd2DSurface
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundPDDIB m_PreviewDIB
        
        PD2D.FillRectangleI cSurface, g_CheckerboardBrush, 0, 0, m_PreviewDIB.GetDIBWidth, m_PreviewDIB.GetDIBHeight
        PD2D.FillRectangleI cSurface, m_Filler, 0, 0, m_PreviewDIB.GetDIBWidth, m_PreviewDIB.GetDIBHeight
        
        'Render a border around the outside
        Dim cPenBorder As pd2DPen
        Set cPenBorder = New pd2DPen
        cPenBorder.SetPenWidth 1
        cPenBorder.SetPenLineJoin P2_LJ_Miter
        If (Not g_Themer Is Nothing) Then cPenBorder.SetPenColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
        PD2D.DrawRectangleI cSurface, cPenBorder, 0, 0, m_PreviewDIB.GetDIBWidth, m_PreviewDIB.GetDIBHeight
        
        'Request a picture box redraw
        Me.picBrushPreview.RequestRedraw True
        
        'Notify our parent of the update
        If (Not m_parentBrushControl Is Nothing) Then m_parentBrushControl.NotifyOfLiveBrushChange m_Filler.GetBrushPropertiesAsXML
        
    End If
    
End Sub

Private Sub gsPrimary_GradientChanged()
    UpdatePreview
End Sub

Private Sub lstFillPattern_Click()
    UpdatePreview
End Sub

Private Sub lstFillPattern_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)
    
    If (bufferDC = 0) Then Exit Sub
    
    Dim tmpRectF As RectF
    CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&
    
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
    
    PD2D.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
    
    'Next, fill the border
    PD2D.DrawRectangleF_FromRectF cSurface, cPen, tmpRectF
    
    'Next, draw the caption
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(10#)
    
    Dim tmpString As String
    tmpString = CStr(itemIndex + 1) & " - "
    
    tmpFont.SetFontColor itemTextColor
    tmpFont.AttachToDC bufferDC
    tmpFont.SetTextAlignment vbLeftJustify
    tmpFont.FastRenderTextWithClipping tmpRectF.Left + FixDPI(4), tmpRectF.Top, tmpRectF.Width, tmpRectF.Height, tmpString, False, True, False
    tmpFont.ReleaseFromDC
    Set tmpFont = Nothing
        
    'Finally, draw the hatch
    Dim hatchRect As RectF
    
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
        
    cSurface.SetSurfaceRenderingOrigin hatchRect.Left, hatchRect.Top
    PD2D.FillRectangleF_FromRectF cSurface, cBrush, hatchRect
    
    cPen.SetPenColor hatchBorderColor
    PD2D.DrawRectangleF_FromRectF cSurface, cPen, hatchRect
    
End Sub

Private Sub picBrushPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    m_PreviewDIB.AlphaBlendToDC targetDC
End Sub

Private Sub sldGradientAngle_Change()
    UpdatePreview
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
        
        btsStyle.ListIndex = .GetBrushMode()
        
        csFillColor.Color = .GetBrushColor()
        sltFillOpacity.Value = .GetBrushOpacity()
        
        lstFillPattern.ListIndex = .GetBrushPatternStyle()
        csPattern(0).Color = .GetBrushPattern1Color()
        csPattern(1).Color = .GetBrushPattern2Color()
        sltPatternOpacity(0).Value = .GetBrushPattern1Opacity()
        sltPatternOpacity(1).Value = .GetBrushPattern2Opacity()
        
        m_Gradient.CreateGradientFromString .GetBrushGradientAllSettings()
        gsPrimary.Gradient = m_Gradient.GetGradientAsString
        btsGradientShape.ListIndex = m_Gradient.GetGradientShape
        sldGradientAngle.Value = m_Gradient.GetGradientAngle
        
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
