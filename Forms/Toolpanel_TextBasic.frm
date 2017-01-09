VERSION 5.00
Begin VB.Form toolpanel_Text 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18465
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
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
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1231
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer picConvertLayer 
      Height          =   1335
      Left            =   17280
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdHyperlink lblConvertLayerConfirm 
         Height          =   240
         Left            =   120
         Top             =   900
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   423
         Alignment       =   2
         Caption         =   "yes"
         Layout          =   2
         RaiseClickEvent =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblConvertLayer 
         Height          =   735
         Left            =   0
         Top             =   120
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   1296
         Alignment       =   2
         Caption         =   "yes"
         Layout          =   1
      End
   End
   Begin PhotoDemon.pdButtonStrip btsHAlignment 
      Height          =   435
      Left            =   15720
      TabIndex        =   9
      Top             =   30
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      ColorScheme     =   1
   End
   Begin PhotoDemon.pdButtonToolbox btnFontStyles 
      Height          =   435
      Index           =   0
      Left            =   7680
      TabIndex        =   5
      Top             =   930
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   767
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltTextClarity 
      Height          =   405
      Left            =   11880
      TabIndex        =   4
      Top             =   930
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   767
      Value           =   5
      NotchPosition   =   2
      NotchValueCustom=   5
   End
   Begin PhotoDemon.pdColorSelector csTextFontColor 
      Height          =   390
      Left            =   11880
      TabIndex        =   0
      Top             =   60
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   688
      curColor        =   0
   End
   Begin PhotoDemon.pdSpinner tudTextFontSize 
      Height          =   345
      Left            =   7680
      TabIndex        =   1
      Top             =   510
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
      DefaultValue    =   16
      Min             =   1
      Max             =   1000
      SigDigits       =   1
      Value           =   16
   End
   Begin PhotoDemon.pdTextBox txtTextTool 
      Height          =   1380
      Left            =   840
      TabIndex        =   2
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
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   3
      Left            =   6360
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
      Left            =   6360
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
      Left            =   6360
      Top             =   1020
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "font style:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdDropDown cboTextRenderingHint 
      Height          =   375
      Left            =   11880
      TabIndex        =   3
      Top             =   525
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   5
      Left            =   10320
      Top             =   570
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "antialiasing:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   6
      Left            =   10320
      Top             =   1020
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "clarity:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   7
      Left            =   10320
      Top             =   120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "color:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdButtonToolbox btnFontStyles 
      Height          =   435
      Index           =   1
      Left            =   8160
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
      Left            =   8640
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
      Left            =   9120
      TabIndex        =   8
      Top             =   930
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   767
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   8
      Left            =   14400
      Top             =   120
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "alignment:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdButtonStrip btsVAlignment 
      Height          =   435
      Left            =   15720
      TabIndex        =   10
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      ColorScheme     =   1
   End
   Begin PhotoDemon.pdDropDownFont cboTextFontFace 
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   60
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   661
   End
End
Attribute VB_Name = "toolpanel_Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Basic Text Tool Panel
'Copyright 2013-2017 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 13/May/15
'Last update: finish migrating all relevant controls to this dedicated form
'
'This form includes all user-editable settings for the Basic Text tool.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

Private Sub btnFontStyles_Click(Index As Integer)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tool_Support.GetToolBusyState)
    If (Not Tool_Support.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.SetToolBusyState True
    
    'Update whichever style was toggled
    Select Case Index
    
        'Bold
        Case 0
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontBold, btnFontStyles(Index).Value
        
        'Italic
        Case 1
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
    'Free the tool engine
    Tool_Support.SetToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub btnFontStyles_GotFocusAPI(Index As Integer)
    
    'Non-destructive effects are obviously not tracked if no images are loaded
    If (g_OpenImageCount = 0) Then Exit Sub
    
    'Set Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            Processor.FlagInitialNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
            
        'Italic
        Case 1
            Processor.FlagInitialNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
        
        'Underline
        Case 2
            Processor.FlagInitialNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
        
        'Strikeout
        Case 3
            Processor.FlagInitialNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
    
    End Select
    
End Sub

Private Sub btnFontStyles_LostFocusAPI(Index As Integer)
    
    If (g_OpenImageCount = 0) Then Exit Sub
    
    'Evaluate Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value
            
        'Italic
        Case 1
            If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
End Sub

Private Sub btsHAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tool_Support.GetToolBusyState)
    If (Not Tool_Support.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_HorizontalAlignment, buttonIndex
    
    'Free the tool engine
    Tool_Support.SetToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub btsHAlignment_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub btsHAlignment_LostFocusAPI()
    If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex
End Sub

Private Sub btsVAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tool_Support.GetToolBusyState)
    If (Not Tool_Support.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_VerticalAlignment, buttonIndex
    
    'Free the tool engine
    Tool_Support.SetToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub btsVAlignment_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub btsVAlignment_LostFocusAPI()
    If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex
End Sub

Private Sub cboTextFontFace_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tool_Support.GetToolBusyState)
    If (Not Tool_Support.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.SetToolBusyState True
    
    'Update the current layer font size
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
    
    'Free the tool engine
    Tool_Support.SetToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboTextFontFace_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex), pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub cboTextFontFace_LostFocusAPI()
    If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
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
        
    'If tool changes are not allowed, exit.  (Note that this also queries Tool_Support.GetToolBusyState)
    If (Not Tool_Support.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.SetToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
    
    'Free the tool engine
    Tool_Support.SetToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboTextRenderingHint_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub cboTextRenderingHint_LostFocusAPI()
    If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
End Sub

Private Sub csTextFontColor_ColorChanged()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tool_Support.GetToolBusyState)
    If (Not Tool_Support.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.SetToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontColor, csTextFontColor.Color
    
    'Free the tool engine
    Tool_Support.SetToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub csTextFontColor_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontColor, csTextFontColor.Color, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub csTextFontColor_LostFocusAPI()
    If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_FontColor, csTextFontColor.Color
End Sub

Private Sub Form_Load()
    
    'Forcibly hide the "convert to text layer" panel
    toolpanel_Text.picConvertLayer.Visible = False
    
    'Generate a list of fonts
    If g_IsProgramRunning Then
        cboTextFontFace.InitializeFontList
        cboTextFontFace.ListIndex = cboTextFontFace.ListIndexByString(g_InterfaceFont, vbBinaryCompare)
    
        cboTextRenderingHint.Clear
        cboTextRenderingHint.AddItem "None", 0
        cboTextRenderingHint.AddItem "Normal", 1
        cboTextRenderingHint.AddItem "Crisp", 2
        cboTextRenderingHint.ListIndex = 1
        
        'Add dummy entries to the various alignment buttons; we'll populate these with theme-specific
        ' images in the UpdateAgainstCurrentTheme() function.
        btsHAlignment.AddItem vbNullString, 0
        btsHAlignment.AddItem vbNullString, 1
        btsHAlignment.AddItem vbNullString, 2
        
        btsVAlignment.AddItem vbNullString, 0
        btsVAlignment.AddItem vbNullString, 1
        btsVAlignment.AddItem vbNullString, 2
        
        'Load any last-used settings for this form
        Set lastUsedSettings = New pdLastUsedSettings
        lastUsedSettings.SetParentForm Me
        lastUsedSettings.LoadAllControlValues
        
        'Update everything against the current theme.  This will also set tooltips for various controls.
        UpdateAgainstCurrentTheme
        
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    If (Not lastUsedSettings Is Nothing) Then
        lastUsedSettings.SaveAllControlValues
        lastUsedSettings.SetParentForm Nothing
    End If
    
End Sub

Private Sub Form_Resize()
    UpdateAgainstCurrentLayer
End Sub

Private Sub lblConvertLayerConfirm_Click()
    
    'Because of the way this warning panel is constructed, this label will not be visible unless a click is valid.
    pdImages(g_CurrentImage).GetActiveLayer.SetLayerType PDL_TEXT
    pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, pdImages(g_CurrentImage).GetActiveLayerIndex
    
    'Hide the warning panel and redraw both the viewport, and the UI (as new UI options may now be available)
    Me.UpdateAgainstCurrentLayer
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    SyncInterfaceToCurrentImage
    
End Sub

Private Sub sltTextClarity_Change()

    'If tool changes are not allowed, exit.  (Note that this also queries Tool_Support.GetToolBusyState)
    If (Not Tool_Support.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.SetToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_TextContrast, sltTextClarity.Value
    
    'Free the tool engine
    Tool_Support.SetToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub sltTextClarity_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_TextContrast, sltTextClarity.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub sltTextClarity_LostFocusAPI()
    If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_TextContrast, sltTextClarity.Value
End Sub

Private Sub tudTextFontSize_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tool_Support.GetToolBusyState)
    If (Not Tool_Support.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.SetToolBusyState True
    
    'Update the current layer font size
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontSize, tudTextFontSize.Value
    
    'Free the tool engine
    Tool_Support.SetToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub tudTextFontSize_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontSize, tudTextFontSize.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub tudTextFontSize_LostFocusAPI()
    If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_FontSize, tudTextFontSize.Value
End Sub

Private Sub txtTextTool_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tool_Support.GetToolBusyState)
    If (Not Tool_Support.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.SetToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_Text, txtTextTool.Text
    
    'Free the tool engine
    Tool_Support.SetToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
End Sub

Private Sub txtTextTool_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_Text, txtTextTool.Text, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub txtTextTool_LostFocusAPI()
    If Tool_Support.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Text ptp_Text, txtTextTool.Text
End Sub

'Outside functions can forcibly request an update against the current layer.  If the current layer is a non-basic-text text layer of
' some type (e.g. typography), an option will be displayed to convert the layer over.
Public Sub UpdateAgainstCurrentLayer()

    If (g_OpenImageCount > 0) Then

        If pdImages(g_CurrentImage).GetActiveLayer.IsLayerText Then
        
            'Check for non-basic-text layers.
            If pdImages(g_CurrentImage).GetActiveLayer.GetLayerType <> PDL_TEXT Then
            
                Select Case pdImages(g_CurrentImage).GetActiveLayer.GetLayerType
                
                    Case PDL_TYPOGRAPHY
                        Dim newMessage As String
                        newMessage = g_Language.TranslateMessage("This layer is a typography layer.  To edit it with the basic text tool, you must first convert it to a basic text layer.")
                        newMessage = newMessage & vbCrLf & g_Language.TranslateMessage("(This action is non-destructive.)")
                        Me.lblConvertLayer.Caption = newMessage
                        
                    'In the future, other text layer types can be added here.
                
                End Select
            
                Me.lblConvertLayerConfirm.Caption = g_Language.TranslateMessage("Click here to convert this layer to a basic text layer.")
                
                'Make the prompt panel the size of the tool window
                Me.picConvertLayer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
                
                'Center all labels on the panel
                Me.lblConvertLayer.SetLeft (Me.ScaleWidth - Me.lblConvertLayer.GetWidth) / 2
                Me.lblConvertLayerConfirm.SetLeft (Me.ScaleWidth - Me.lblConvertLayerConfirm.GetWidth) / 2
                
                'Display the panel
                Me.picConvertLayer.Visible = True
                
            Else
                Me.picConvertLayer.Visible = False
            End If
        
        Else
            Me.picConvertLayer.Visible = False
        End If
        
    Else
        Me.picConvertLayer.Visible = False
    End If

End Sub

'Most objects on this form can avoid doing any work if the current layer is not a text layer.
Private Function CurrentLayerIsText() As Boolean
    
    CurrentLayerIsText = False
    
    'Changing UI elements does nothing if no images are loaded
    If (g_OpenImageCount = 0) Then Exit Function
    
    If (Not pdImages(g_CurrentImage) Is Nothing) Then
        If (Not pdImages(g_CurrentImage).GetActiveLayer Is Nothing) Then
            CurrentLayerIsText = pdImages(g_CurrentImage).GetActiveLayer.IsLayerText
        End If
    End If
    
End Function

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Update any UI images against the current theme
    Dim buttonSize As Long
    buttonSize = FixDPI(24)
    
    btnFontStyles(0).AssignImage "format_bold", , buttonSize, buttonSize
    btnFontStyles(1).AssignImage "format_italic", , buttonSize, buttonSize
    btnFontStyles(2).AssignImage "format_underline", , buttonSize, buttonSize
    btnFontStyles(3).AssignImage "format_strikethrough", , buttonSize, buttonSize
    
    btsHAlignment.AssignImageToItem 0, "format_alignleft", , buttonSize, buttonSize
    btsHAlignment.AssignImageToItem 1, "format_aligncenter", , buttonSize, buttonSize
    btsHAlignment.AssignImageToItem 2, "format_alignright", , buttonSize, buttonSize
    
    btsVAlignment.AssignImageToItem 0, "format_aligntop", , buttonSize, buttonSize
    btsVAlignment.AssignImageToItem 1, "format_alignmiddle", , buttonSize, buttonSize
    btsVAlignment.AssignImageToItem 2, "format_alignbottom", , buttonSize, buttonSize
        
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me

End Sub

