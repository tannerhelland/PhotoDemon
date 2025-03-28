VERSION 5.00
Begin VB.Form options_Interface 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   ControlBox      =   0   'False
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
   Icon            =   "Options_Interface.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdCheckBox chkZoomMouse 
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   1965
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   635
      Caption         =   "zoom with mouse wheel"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdSpinner spnSnapDistance 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      DefaultValue    =   8
      Min             =   1
      Max             =   255
      Value           =   8
   End
   Begin PhotoDemon.pdPictureBox picGrid 
      Height          =   735
      Left            =   150
      Top             =   5760
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdColorSelector csCanvasColor 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   375
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdButtonStrip btsTitleText 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1720
      Caption         =   "title bar text:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdLabel lblRecentFileCount 
      Height          =   240
      Left            =   135
      Top             =   2820
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   423
      Caption         =   "maximum number of recent files to remember: "
      ForeColor       =   4210752
      Layout          =   2
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   13
      Left            =   0
      Top             =   2415
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   503
      Caption         =   "recent files"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   14
      Left            =   0
      Top             =   0
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   503
      Caption         =   "main window"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdDropDown cboAlphaCheckSize 
      Height          =   810
      Left            =   1080
      TabIndex        =   4
      Top             =   5700
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1429
      Caption         =   "grid size:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboAlphaCheck 
      Height          =   795
      Left            =   4140
      TabIndex        =   5
      Top             =   5700
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1402
      Caption         =   "grid colors:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdColorSelector csAlphaOne 
      Height          =   690
      Left            =   7260
      TabIndex        =   6
      Top             =   5760
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   1217
      ShowMainWindowColor=   0   'False
   End
   Begin PhotoDemon.pdColorSelector csAlphaTwo 
      Height          =   690
      Left            =   7770
      TabIndex        =   7
      Top             =   5760
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   1217
      ShowMainWindowColor=   0   'False
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   2
      Left            =   0
      Top             =   5280
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   503
      Caption         =   "transparency"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdButtonStrip btsMRUStyle 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   3210
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1508
      Caption         =   "recent file menu text:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdLabel lblCanvasColor 
      Height          =   240
      Left            =   120
      Top             =   420
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   423
      Caption         =   "canvas background color:"
      ForeColor       =   4210752
      Layout          =   2
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   23
      Left            =   0
      Top             =   4320
      Width           =   4020
      _ExtentX        =   14288
      _ExtentY        =   503
      Caption         =   "snap distance (in pixels)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   25
      Left            =   4080
      Top             =   4320
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   503
      Caption         =   "angle snap distance (in degrees)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdSpinner spnSnapDistance 
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   9
      Top             =   4800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      DefaultValue    =   5
      Min             =   1
      Max             =   15
      SigDigits       =   1
      Value           =   5
   End
   Begin PhotoDemon.pdSpinner tudRecentFiles 
      Height          =   345
      Left            =   3840
      TabIndex        =   10
      Top             =   2760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      DefaultValue    =   10
      Min             =   1
      Max             =   32
      Value           =   10
   End
End
Attribute VB_Name = "options_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Used to see if the user physically clicked a combo box, or if VB selected it on its own
Private m_userInitiatedAlphaSelection As Boolean

'Alpha channel checkerboard selection; change the color selectors to match
Private Sub cboAlphaCheck_Click()

    'Only respond to user-generated events (e.g. do *not* trigger during form initialization)
    If m_userInitiatedAlphaSelection Then

        m_userInitiatedAlphaSelection = False

        'Redraw the sample picture boxes based on the value the user has selected
        Select Case cboAlphaCheck.ListIndex
        
            'highlights
            Case 0
                csAlphaOne.Color = RGB(255, 255, 255)
                csAlphaTwo.Color = RGB(204, 204, 204)
            
            'midtones
            Case 1
                csAlphaOne.Color = RGB(153, 153, 153)
                csAlphaTwo.Color = RGB(102, 102, 102)
            
            'shadows
            Case 2
                csAlphaOne.Color = RGB(51, 51, 51)
                csAlphaTwo.Color = RGB(0, 0, 0)
            
            'red
            Case 3
                csAlphaOne.Color = RGB(255, 255, 255)
                csAlphaTwo.Color = RGB(255, 200, 200)
            
            'orange
            Case 4
                csAlphaOne.Color = RGB(255, 255, 255)
                csAlphaTwo.Color = RGB(255, 215, 170)
            
            'green
            Case 5
                csAlphaOne.Color = RGB(255, 255, 255)
                csAlphaTwo.Color = RGB(200, 255, 200)
            
            'blue
            Case 6
                csAlphaOne.Color = RGB(255, 255, 255)
                csAlphaTwo.Color = RGB(200, 225, 255)
            
            'purple
            Case 7
                csAlphaOne.Color = RGB(255, 255, 255)
                csAlphaTwo.Color = RGB(225, 200, 255)
            
            'custom
            Case 8
                csAlphaOne.Color = RGB(255, 160, 60)
                csAlphaTwo.Color = RGB(160, 240, 160)
            
        End Select
        
        'Redraw the "sample" grid
        picGrid.RequestRedraw True
        
        m_userInitiatedAlphaSelection = True
                
    End If
    
    UpdateAlphaGridVisibility
    
End Sub

Private Sub cboAlphaCheckSize_Click()
    picGrid.RequestRedraw True
End Sub

'When new transparency checkerboard colors are selected, change the corresponding list box to match
Private Sub csAlphaOne_ColorChanged()
    
    If m_userInitiatedAlphaSelection Then
        m_userInitiatedAlphaSelection = False
        picGrid.RequestRedraw True
        cboAlphaCheck.ListIndex = 8     '"custom colors"
        m_userInitiatedAlphaSelection = True
    End If
    
End Sub

Private Sub csAlphaTwo_ColorChanged()
    
    If m_userInitiatedAlphaSelection Then
        picGrid.RequestRedraw
        m_userInitiatedAlphaSelection = False
        cboAlphaCheck.ListIndex = 8     '"custom colors"
        m_userInitiatedAlphaSelection = True
    End If
    
End Sub

Private Sub Form_Load()

    'Interface prefs
    btsTitleText.AddItem "compact (filename only)", 0
    btsTitleText.AddItem "verbose (filename and path)", 1
    btsTitleText.AssignTooltip "The title bar of the main PhotoDemon window displays information about the currently loaded image.  Use this preference to control how much information is displayed."
    
    lblCanvasColor.Caption = g_Language.TranslateMessage("canvas background color: ")
    csCanvasColor.SetLeft lblCanvasColor.GetLeft + lblCanvasColor.GetWidth + Interface.FixDPI(8)
    csCanvasColor.SetWidth (btsTitleText.GetLeft + btsTitleText.GetWidth) - (csCanvasColor.GetLeft)
    
    lblRecentFileCount.Caption = g_Language.TranslateMessage("maximum number of recent files to remember: ")
    tudRecentFiles.SetLeft lblRecentFileCount.GetLeft + lblRecentFileCount.GetWidth + Interface.FixDPI(8)
    
    btsMRUStyle.AddItem "compact (filename only)", 0
    btsMRUStyle.AddItem "verbose (filename and path)", 1
    btsMRUStyle.AssignTooltip "The ""Recent Files"" menu width is limited by Windows.  To prevent this menu from overflowing, PhotoDemon can display image names only instead of full image locations."
    
    chkZoomMouse.Value = UserPrefs.GetPref_Boolean("Interface", "wheel-zoom", False)
    
    m_userInitiatedAlphaSelection = False
    cboAlphaCheck.Clear
    cboAlphaCheck.AddItem "highlights", 0
    cboAlphaCheck.AddItem "midtones", 1
    cboAlphaCheck.AddItem "shadows", 2, True
    cboAlphaCheck.AddItem "red", 3
    cboAlphaCheck.AddItem "orange", 4
    cboAlphaCheck.AddItem "green", 5
    cboAlphaCheck.AddItem "blue", 6
    cboAlphaCheck.AddItem "purple", 7, True
    cboAlphaCheck.AddItem "custom", 8
    cboAlphaCheck.AssignTooltip "To help identify transparent pixels, a special grid appears ""behind"" them.  This setting modifies the grid's appearance."
    m_userInitiatedAlphaSelection = True
    
    cboAlphaCheckSize.Clear
    cboAlphaCheckSize.AddItem "small", 0
    cboAlphaCheckSize.AddItem "medium", 1
    cboAlphaCheckSize.AddItem "large", 2
    cboAlphaCheckSize.AssignTooltip "To help identify transparent pixels, a special grid appears ""behind"" them.  This setting modifies the grid's appearance."
    
End Sub

Private Sub picGrid_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    
    Dim chkSize As Long
    Select Case cboAlphaCheckSize.ListIndex
        Case 0
            chkSize = 4
        Case 1
            chkSize = 8
        Case 2
            chkSize = 16
    End Select
    
    Dim tmpGrid As pdDIB
    Set tmpGrid = New pdDIB
    Drawing.GetArbitraryCheckerboardDIB tmpGrid, csAlphaOne.Color, csAlphaTwo.Color, chkSize
    
    Dim tmpBrush As pd2DBrush
    Set tmpBrush = New pd2DBrush
    tmpBrush.SetBrushMode P2_BM_Texture
    tmpBrush.SetBrushTextureWrapMode P2_WM_Tile
    tmpBrush.SetBrushTextureFromDIB tmpGrid
    
    Dim tmpSurface As pd2DSurface
    Set tmpSurface = New pd2DSurface
    tmpSurface.WrapSurfaceAroundDC targetDC
    tmpSurface.SetSurfaceAntialiasing P2_AA_None
    tmpSurface.SetSurfacePixelOffset P2_PO_Normal
    tmpSurface.SetSurfaceRenderingOrigin 1, 1
    
    PD2D.FillRectangleI tmpSurface, tmpBrush, 0, 0, ctlWidth, ctlHeight
    
    Dim tmpPen As pd2DPen
    Drawing2D.QuickCreateSolidPen tmpPen, 1, g_Themer.GetGenericUIColor(UI_GrayNeutral)
    PD2D.DrawRectangleI tmpSurface, tmpPen, 0, 0, ctlWidth - 1, ctlHeight - 1
    
End Sub

Public Sub LoadUserPreferences()

    'Interface preferences
    btsTitleText.ListIndex = UserPrefs.GetPref_Long("Interface", "Window Caption Length", 0)
    csCanvasColor.Color = UserPrefs.GetCanvasColor()
    tudRecentFiles.Value = UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)
    btsMRUStyle.ListIndex = UserPrefs.GetPref_Long("Interface", "MRU Caption Length", 0)
    spnSnapDistance(0).Value = UserPrefs.GetPref_Long("Interface", "snap-distance", 8&)
    spnSnapDistance(1).Value = UserPrefs.GetPref_Float("Interface", "snap-degrees", 7.5)
    m_userInitiatedAlphaSelection = False
    cboAlphaCheck.ListIndex = UserPrefs.GetPref_Long("Transparency", "Alpha Check Mode", 0)
    csAlphaOne.Color = UserPrefs.GetPref_Long("Transparency", "Alpha Check One", RGB(255, 255, 255))
    csAlphaTwo.Color = UserPrefs.GetPref_Long("Transparency", "Alpha Check Two", RGB(204, 204, 204))
    m_userInitiatedAlphaSelection = True
    cboAlphaCheckSize.ListIndex = UserPrefs.GetPref_Long("Transparency", "Alpha Check Size", 1)
    UpdateAlphaGridVisibility
    
End Sub

Public Sub SaveUserPreferences()

    'Interface preferences
    UserPrefs.SetPref_Long "Interface", "Window Caption Length", btsTitleText.ListIndex
    UserPrefs.SetPref_String "Interface", "Canvas Color", Colors.GetHexStringFromRGB(csCanvasColor.Color)
    UserPrefs.SetCanvasColor csCanvasColor.Color
    
    'Changes to the recent files list (including count and how it's displayed) may require us to
    ' trigger a full rebuild of the menu
    Dim mruNeedsToBeRebuilt As Boolean
    mruNeedsToBeRebuilt = (btsMRUStyle.ListIndex <> UserPrefs.GetPref_Long("Interface", "MRU Caption Length", 0))
    UserPrefs.SetPref_Long "Interface", "MRU Caption Length", btsMRUStyle.ListIndex
    
    Dim newMaxRecentFiles As Long
    If tudRecentFiles.IsValid Then newMaxRecentFiles = tudRecentFiles.Value Else newMaxRecentFiles = 10
    If (Not mruNeedsToBeRebuilt) Then mruNeedsToBeRebuilt = (newMaxRecentFiles <> UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10))
    UserPrefs.SetPref_Long "Interface", "Recent Files Limit", tudRecentFiles.Value
    
    'If any MRUs need to be rebuilt, do so now
    If mruNeedsToBeRebuilt Then
        g_RecentFiles.NotifyMaxLimitChanged
        g_RecentMacros.MRU_NotifyNewMaxLimit
    End If
    
    UserPrefs.SetPref_Boolean "Interface", "wheel-zoom", chkZoomMouse.Value
    UserPrefs.SetZoomWithWheel chkZoomMouse.Value
    
    UserPrefs.SetPref_Long "Interface", "snap-distance", spnSnapDistance(0).Value
    UserPrefs.SetPref_Long "Interface", "snap-degrees", spnSnapDistance(1).Value
    Snap.SetSnap_Distance spnSnapDistance(0).Value
    Snap.SetSnap_Degrees spnSnapDistance(1).Value
    
    UserPrefs.SetPref_Long "Transparency", "Alpha Check Mode", CLng(cboAlphaCheck.ListIndex)
    UserPrefs.SetPref_Long "Transparency", "Alpha Check One", CLng(csAlphaOne.Color)
    UserPrefs.SetPref_Long "Transparency", "Alpha Check Two", CLng(csAlphaTwo.Color)
    UserPrefs.SetPref_Long "Transparency", "Alpha Check Size", cboAlphaCheckSize.ListIndex
    Drawing.CreateAlphaCheckerboardDIB g_CheckerboardPattern
    
End Sub

Private Sub UpdateAlphaGridVisibility()
    Dim colorBoxVisibility As Boolean
    colorBoxVisibility = (cboAlphaCheck.ListIndex = 8)
    csAlphaOne.Visible = colorBoxVisibility
    csAlphaTwo.Visible = colorBoxVisibility
End Sub
