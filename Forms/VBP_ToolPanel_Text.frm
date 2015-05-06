VERSION 5.00
Begin VB.Form toolpanel_Text 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18465
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
   ScaleWidth      =   1231
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.buttonStrip btsHAlignment 
      Height          =   435
      Left            =   15720
      TabIndex        =   10
      Top             =   30
      Width           =   1455
      _extentx        =   2566
      _extenty        =   767
      colorscheme     =   1
   End
   Begin PhotoDemon.pdButtonToolbox btnFontStyles 
      Height          =   435
      Index           =   0
      Left            =   7680
      TabIndex        =   6
      Top             =   930
      Width           =   450
      _extentx        =   794
      _extenty        =   767
      stickytoggle    =   -1
   End
   Begin PhotoDemon.sliderTextCombo sltTextClarity 
      Height          =   435
      Left            =   11880
      TabIndex        =   5
      Top             =   930
      Width           =   2415
      _extentx        =   4260
      _extenty        =   767
      value           =   5
      notchposition   =   2
      notchvaluecustom=   5
   End
   Begin PhotoDemon.colorSelector csTextFontColor 
      Height          =   390
      Left            =   11880
      TabIndex        =   0
      Top             =   60
      Width           =   2415
      _extentx        =   4260
      _extenty        =   688
      curcolor        =   0
   End
   Begin PhotoDemon.textUpDown tudTextFontSize 
      Height          =   345
      Left            =   7680
      TabIndex        =   1
      Top             =   510
      Width           =   2415
      _extentx        =   4260
      _extenty        =   609
      max             =   1000
      min             =   1
      value           =   16
   End
   Begin PhotoDemon.pdTextBox txtTextTool 
      Height          =   1380
      Left            =   840
      TabIndex        =   2
      Top             =   30
      Width           =   5295
      _extentx        =   9340
      _extenty        =   2434
      fontsize        =   9
      multiline       =   -1
   End
   Begin PhotoDemon.pdComboBox cboTextFontFace 
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   60
      Width           =   2415
      _extentx        =   4260
      _extenty        =   635
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   1080
      Index           =   0
      Left            =   0
      Top             =   360
      Width           =   765
      _extentx        =   1349
      _extenty        =   1905
      alignment       =   2
      caption         =   "(this tool is under constr- uction)"
      forecolor       =   255
      layout          =   1
      usecustomforecolor=   -1
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   1
      Left            =   120
      Top             =   60
      Width           =   645
      _extentx        =   1138
      _extenty        =   503
      alignment       =   1
      caption         =   "text:"
      forecolor       =   0
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   3
      Left            =   6360
      Top             =   120
      Width           =   1125
      _extentx        =   1984
      _extenty        =   503
      alignment       =   1
      caption         =   "font face:"
      forecolor       =   0
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   4
      Left            =   6360
      Top             =   570
      Width           =   1125
      _extentx        =   1984
      _extenty        =   503
      alignment       =   1
      caption         =   "font size:"
      forecolor       =   0
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   2
      Left            =   6360
      Top             =   1020
      Width           =   1125
      _extentx        =   1984
      _extenty        =   503
      alignment       =   1
      caption         =   "font style:"
      forecolor       =   0
   End
   Begin PhotoDemon.pdComboBox cboTextRenderingHint 
      Height          =   375
      Left            =   11880
      TabIndex        =   4
      Top             =   525
      Width           =   2415
      _extentx        =   4260
      _extenty        =   635
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   5
      Left            =   10320
      Top             =   570
      Width           =   1365
      _extentx        =   2408
      _extenty        =   503
      alignment       =   1
      caption         =   "antialiasing:"
      forecolor       =   0
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   6
      Left            =   10320
      Top             =   1020
      Width           =   1365
      _extentx        =   2408
      _extenty        =   503
      alignment       =   1
      caption         =   "clarity:"
      forecolor       =   0
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   7
      Left            =   10320
      Top             =   120
      Width           =   1365
      _extentx        =   2408
      _extenty        =   503
      alignment       =   1
      caption         =   "color:"
      forecolor       =   0
   End
   Begin PhotoDemon.pdButtonToolbox btnFontStyles 
      Height          =   435
      Index           =   1
      Left            =   8160
      TabIndex        =   7
      Top             =   930
      Width           =   450
      _extentx        =   794
      _extenty        =   767
      stickytoggle    =   -1
   End
   Begin PhotoDemon.pdButtonToolbox btnFontStyles 
      Height          =   435
      Index           =   2
      Left            =   8640
      TabIndex        =   8
      Top             =   930
      Width           =   450
      _extentx        =   794
      _extenty        =   767
      stickytoggle    =   -1
   End
   Begin PhotoDemon.pdButtonToolbox btnFontStyles 
      Height          =   435
      Index           =   3
      Left            =   9120
      TabIndex        =   9
      Top             =   930
      Width           =   450
      _extentx        =   794
      _extenty        =   767
      stickytoggle    =   -1
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   8
      Left            =   14400
      Top             =   120
      Width           =   1125
      _extentx        =   1984
      _extenty        =   503
      alignment       =   1
      caption         =   "alignment:"
      forecolor       =   0
   End
   Begin PhotoDemon.buttonStrip btsVAlignment 
      Height          =   435
      Left            =   15720
      TabIndex        =   11
      Top             =   480
      Width           =   1455
      _extentx        =   2566
      _extenty        =   767
      colorscheme     =   1
   End
End
Attribute VB_Name = "toolpanel_Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    Processor.flagInitialNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex), pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboTextFontFace_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
End Sub

Private Sub cboTextRenderingHint_Click()
        
    'We show/hide the AA clarity option depending on this tool's setting.  (AA clarity doesn't make much sense
    ' if AA is disabled.)
    If cboTextRenderingHint.ListIndex = 0 Then
        sltTextClarity.Visible = False
        lblText(6).Visible = False
    Else
        sltTextClarity.Visible = True
        lblText(6).Visible = True
    End If
        
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
    Processor.flagInitialNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboTextRenderingHint_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
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

Private Sub Form_Load()

    'Generate a list of fonts
    If g_IsProgramRunning Then
    
        'Retrieve a copy of the current system font cache
        Font_Management.getCopyOfFontCache userFontList
        
        'Populate the font selection combo box
        Dim tmpFontName As String, relevantListIndex As Long
        
        Dim i As Long
        For i = 0 To userFontList.getNumOfStrings - 1
            cboTextFontFace.AddItem userFontList.GetString(i)
            If StrComp(userFontList.GetString(i), g_InterfaceFont) = 0 Then relevantListIndex = i
        Next i
        
        cboTextFontFace.ListIndex = relevantListIndex
        
    End If
    
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

Private Sub sltTextClarity_Change()

    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_TextContrast, sltTextClarity.Value
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

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
