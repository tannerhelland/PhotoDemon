VERSION 5.00
Begin VB.Form toolpanel_FancyText 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1153
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.buttonStripVertical btsCategory 
      Height          =   1380
      Left            =   6240
      TabIndex        =   1
      Top             =   30
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2434
   End
   Begin PhotoDemon.pdTextBox txtTextTool 
      Height          =   1380
      Left            =   840
      TabIndex        =   0
      Top             =   30
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2434
      FontSize        =   9
      Multiline       =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   1
      Left            =   120
      Top             =   60
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "text:"
      ForeColor       =   0
   End
   Begin VB.PictureBox picCategory 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   0
      Left            =   8640
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   2
      Top             =   0
      Width           =   10935
      Begin PhotoDemon.pdComboBox_Font cboTextFontFace 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   60
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   930
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.textUpDown tudTextFontSize 
         Height          =   345
         Left            =   1320
         TabIndex        =   5
         Top             =   510
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
         Min             =   1
         Max             =   1000
         Value           =   16
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   3
         Left            =   0
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "font face:"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   4
         Left            =   0
         Top             =   570
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "font size:"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   2
         Left            =   0
         Top             =   1020
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "font style:"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   1
         Left            =   1800
         TabIndex        =   6
         Top             =   930
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   2
         Left            =   2280
         TabIndex        =   7
         Top             =   930
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   3
         Left            =   2760
         TabIndex        =   8
         Top             =   930
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.smartCheckBox chkHinting 
         Height          =   330
         Left            =   4200
         TabIndex        =   9
         Top             =   510
         Width           =   1815
         _ExtentX        =   2990
         _ExtentY        =   582
         Caption         =   "hinting"
         Value           =   0
      End
      Begin PhotoDemon.pdComboBox cboTextRenderingHint 
         Height          =   375
         Left            =   5400
         TabIndex        =   15
         Top             =   60
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   5
         Left            =   3840
         Top             =   120
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "antialiasing:"
         ForeColor       =   0
      End
   End
   Begin VB.PictureBox picCategory 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   2
      Left            =   8640
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   10935
      Begin PhotoDemon.colorSelector csTextFontColor 
         Height          =   390
         Left            =   1320
         TabIndex        =   14
         Top             =   60
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   688
         curColor        =   0
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   7
         Left            =   0
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "color:"
         ForeColor       =   0
      End
   End
   Begin VB.PictureBox picCategory 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   1
      Left            =   8640
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   10935
      Begin PhotoDemon.buttonStrip btsHAlignment 
         Height          =   435
         Left            =   1320
         TabIndex        =   11
         Top             =   60
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         ColorScheme     =   1
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   8
         Left            =   0
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "alignment:"
         ForeColor       =   0
      End
      Begin PhotoDemon.buttonStrip btsVAlignment 
         Height          =   435
         Left            =   1320
         TabIndex        =   12
         Top             =   510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         ColorScheme     =   1
      End
   End
End
Attribute VB_Name = "toolpanel_FancyText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Advanced Typography Tool Panel
'Copyright 2013-2015 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 13/May/15
'Last update: finish migrating all relevant controls to this dedicated form
'
'This form includes all user-editable settings for PD's Advanced Typography text tool.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'Current list of fonts, in pdStringStack format
Private userFontList As pdStringStack

Private Sub btnFontStyles_Click(Index As Integer)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update whichever style was toggled
    Select Case Index
    
        'Bold
        Case 0
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontBold, btnFontStyles(Index).Value
        
        'Italic
        Case 1
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub btnFontStyles_GotFocusAPI(Index As Integer)
    
    'Non-destructive effects are obviously not tracked if no images are loaded
    If g_OpenImageCount = 0 Then Exit Sub
    
    'Set Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            Processor.flagInitialNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
            
        'Italic
        Case 1
            Processor.flagInitialNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
        
        'Underline
        Case 2
            Processor.flagInitialNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
        
        'Strikeout
        Case 3
            Processor.flagInitialNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
    
    End Select
    
End Sub

Private Sub btnFontStyles_LostFocusAPI(Index As Integer)
    
    'Evaluate Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value
            
        'Italic
        Case 1
            If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
End Sub

Private Sub btsCategory_Click(ByVal buttonIndex As Long)
    
    'When the current category is changed, show the relevant panel and hide all others
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        picCategory(i).Visible = CBool(i = buttonIndex)
    Next i
    
End Sub

Private Sub btsHAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_HorizontalAlignment, buttonIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub btsHAlignment_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub btsHAlignment_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex
End Sub

Private Sub btsVAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_VerticalAlignment, buttonIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub btsVAlignment_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub btsVAlignment_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex
End Sub

Private Sub cboTextFontFace_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer font size
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboTextFontFace_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex), pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboTextFontFace_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
End Sub

Private Sub cboTextRenderingHint_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboTextRenderingHint_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboTextRenderingHint_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
End Sub

Private Sub chkHinting_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_TextHinting, CBool(chkHinting.Value)
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub chkHinting_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_TextHinting, CBool(chkHinting.Value), pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub chkHinting_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_TextHinting, CBool(chkHinting.Value)
End Sub

Private Sub csTextFontColor_ColorChanged()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontColor, csTextFontColor.Color
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub csTextFontColor_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_FontColor, csTextFontColor.Color, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub csTextFontColor_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontColor, csTextFontColor.Color
End Sub

Private Sub Form_Load()

    'Generate a list of fonts
    If g_IsProgramRunning Then
        
        'Initialize the font list
        cboTextFontFace.initializeFontList
        
        'Set the system font as the default
        cboTextFontFace.setListIndexByString g_InterfaceFont, vbBinaryCompare
        
    End If
    
    'Draw the category selector
    btsCategory.AddItem "character", 0
    btsCategory.AddItem "paragraph", 1
    btsCategory.AddItem "color", 2
    
    'Fill AA options
    cboTextRenderingHint.Clear
    cboTextRenderingHint.AddItem "None", 0
    cboTextRenderingHint.AddItem "Normal", 1
    cboTextRenderingHint.AddItem "Crisp", 2
    cboTextRenderingHint.ListIndex = 1
    
    'Draw font style buttons
    btnFontStyles(0).AssignImage "TEXT_BOLD"
    btnFontStyles(1).AssignImage "TEXT_ITALIC"
    btnFontStyles(2).AssignImage "TEXT_UNDERLINE"
    btnFontStyles(3).AssignImage "TEXT_STRIKE"
    
    'Draw alignment buttons
    btsHAlignment.AddItem "", 0
    btsHAlignment.AddItem "", 1
    btsHAlignment.AddItem "", 2
    
    btsHAlignment.AssignImageToItem 0, "TEXT_ALIGN_LEFT"
    btsHAlignment.AssignImageToItem 1, "TEXT_ALIGN_HCENTER"
    btsHAlignment.AssignImageToItem 2, "TEXT_ALIGN_RIGHT"
    
    btsVAlignment.AddItem "", 0
    btsVAlignment.AddItem "", 1
    btsVAlignment.AddItem "", 2
    
    btsVAlignment.AssignImageToItem 0, "TEXT_ALIGN_TOP"
    btsVAlignment.AssignImageToItem 1, "TEXT_ALIGN_VCENTER"
    btsVAlignment.AssignImageToItem 2, "TEXT_ALIGN_BOTTOM"
        
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.setParentForm Me
    lastUsedSettings.loadAllControlValues
    
    'Update everything against the current theme.  This will also set tooltips for various controls.
    updateAgainstCurrentTheme

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    lastUsedSettings.saveAllControlValues
    lastUsedSettings.setParentForm Nothing
    
End Sub

Private Sub lastUsedSettings_ReadCustomPresetData()

    'Make sure the right panel is shown
    btsCategory_Click btsCategory.ListIndex

End Sub

Private Sub tudTextFontSize_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer font size
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontSize, tudTextFontSize.Value
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub tudTextFontSize_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_FontSize, tudTextFontSize.Value, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub tudTextFontSize_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontSize, tudTextFontSize.Value
End Sub

Private Sub txtTextTool_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_Text, txtTextTool.Text
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
End Sub

Private Sub txtTextTool_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_Text, txtTextTool.Text, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub txtTextTool_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_Text, txtTextTool.Text
End Sub


'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) MakeFormPretty is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub updateAgainstCurrentTheme()

    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    makeFormPretty Me

End Sub
