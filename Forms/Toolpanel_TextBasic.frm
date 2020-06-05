VERSION 5.00
Begin VB.Form toolpanel_TextBasic 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18465
   ControlBox      =   0   'False
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
   Icon            =   "Toolpanel_TextBasic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1231
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdButtonStripVertical btsMain 
      Height          =   1350
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   2381
   End
   Begin PhotoDemon.pdContainer picConvertLayer 
      Height          =   1335
      Left            =   17280
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
   Begin PhotoDemon.pdContainer pdcMain 
      Height          =   1500
      Index           =   1
      Left            =   2280
      Top             =   0
      Width           =   10980
      _ExtentX        =   18521
      _ExtentY        =   2646
      Begin PhotoDemon.pdSlider sldTextFontSize 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   465
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         Min             =   1
         Max             =   1000
         ScaleStyle      =   2
         Value           =   16
         NotchPosition   =   2
         NotchValueCustom=   16
      End
      Begin PhotoDemon.pdButtonStrip btsHAlignment 
         Height          =   435
         Left            =   9360
         TabIndex        =   5
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         ColorScheme     =   1
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   0
         Left            =   1320
         TabIndex        =   6
         Top             =   915
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSlider sltTextClarity 
         Height          =   405
         Left            =   5520
         TabIndex        =   7
         Top             =   900
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         Value           =   5
         NotchPosition   =   2
         NotchValueCustom=   5
      End
      Begin PhotoDemon.pdColorSelector csTextFontColor 
         Height          =   390
         Left            =   5520
         TabIndex        =   8
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   688
         curColor        =   0
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   3
         Left            =   0
         Top             =   90
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
         Top             =   510
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
         Top             =   975
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "font style:"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdDropDown cboTextRenderingHint 
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         Top             =   495
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   5
         Left            =   3960
         Top             =   540
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
         Left            =   3960
         Top             =   990
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
         Left            =   3960
         Top             =   90
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
         Left            =   1800
         TabIndex        =   10
         Top             =   915
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   2
         Left            =   2280
         TabIndex        =   11
         Top             =   915
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   3
         Left            =   2760
         TabIndex        =   12
         Top             =   915
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   8
         Left            =   8040
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "alignment:"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdButtonStrip btsVAlignment 
         Height          =   435
         Left            =   9360
         TabIndex        =   2
         Top             =   450
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         ColorScheme     =   1
      End
      Begin PhotoDemon.pdDropDownFont cboTextFontFace 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   45
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   661
      End
   End
   Begin PhotoDemon.pdContainer pdcMain 
      Height          =   1500
      Index           =   0
      Left            =   2280
      Top             =   0
      Width           =   10980
      _ExtentX        =   18521
      _ExtentY        =   2646
      Begin PhotoDemon.pdTextBox txtTextTool 
         Height          =   1350
         Left            =   0
         TabIndex        =   4
         Top             =   30
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   2381
         FontSize        =   9
         Multiline       =   -1  'True
      End
   End
End
Attribute VB_Name = "toolpanel_TextBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Basic Text Tool Panel
'Copyright 2013-2020 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 22/June/17
'Last update: large improvements to the way non-destructive actions interact with the Undo/Redo engine
'
'This form includes all user-editable settings for the Basic Text tool.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

Private Sub btnFontStyles_Click(Index As Integer)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update whichever style was toggled
    Select Case Index
    
        'Bold
        Case 0
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontBold, btnFontStyles(Index).Value
        
        'Italic
        Case 1
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

Private Sub btnFontStyles_GotFocusAPI(Index As Integer)
    
    'Non-destructive effects are obviously not tracked if no images are loaded
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    'Set Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            Processor.FlagInitialNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
            
        'Italic
        Case 1
            Processor.FlagInitialNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
        
        'Underline
        Case 2
            Processor.FlagInitialNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
        
        'Strikeout
        Case 3
            Processor.FlagInitialNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
    
    End Select
    
End Sub

Private Sub btnFontStyles_LostFocusAPI(Index As Integer)
    
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    'Evaluate Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            Processor.FlagFinalNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value
            
        'Italic
        Case 1
            Processor.FlagFinalNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            Processor.FlagFinalNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            Processor.FlagFinalNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
End Sub

Private Sub btsHAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_HorizontalAlignment, buttonIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub btsHAlignment_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub btsHAlignment_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex
End Sub

Private Sub btsMain_Click(ByVal buttonIndex As Long)
    ChangeMainPanel
End Sub

Private Sub btsVAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_VerticalAlignment, buttonIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub btsVAlignment_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub btsVAlignment_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex
End Sub

Private Sub cboTextFontFace_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer font size
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub cboTextFontFace_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex), PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub cboTextFontFace_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
End Sub

Private Sub cboTextRenderingHint_Click()
        
    'We show/hide the AA clarity option depending on this tool's setting.  (AA clarity doesn't make much sense
    ' if AA is disabled.)
    If (cboTextRenderingHint.ListIndex = 0) Then
        sltTextClarity.Visible = False
        lblText(6).Visible = False
    Else
        sltTextClarity.Visible = True
        lblText(6).Visible = True
    End If
        
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub cboTextRenderingHint_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub cboTextRenderingHint_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
End Sub

Private Sub csTextFontColor_ColorChanged()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontColor, csTextFontColor.Color
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub csTextFontColor_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontColor, csTextFontColor.Color, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub csTextFontColor_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FontColor, csTextFontColor.Color
End Sub

Private Sub Form_Load()
    
    'Disable any layer updates as a result of control changes during the load process
    Tools.SetToolBusyState True
    
    'Forcibly hide the "convert to text layer" panel
    toolpanel_TextBasic.picConvertLayer.Visible = False
    
    If PDMain.IsProgramRunning() Then
        
        'This tool is separated into two panels: text entry, and text settings
        btsMain.AddItem "text", 0
        btsMain.AddItem "settings", 1
        btsMain.ListIndex = 0
        ChangeMainPanel
        
        'Generate a list of fonts
        cboTextFontFace.InitializeFontList
        cboTextFontFace.ListIndex = cboTextFontFace.ListIndexByString(Fonts.GetUIFontName(), vbBinaryCompare)
        
        'Antialiasing options behave slightly differently from the advanced text tool
        cboTextRenderingHint.SetAutomaticRedraws False
        cboTextRenderingHint.Clear
        cboTextRenderingHint.AddItem "None", 0
        cboTextRenderingHint.AddItem "Normal", 1
        cboTextRenderingHint.AddItem "Crisp", 2
        cboTextRenderingHint.ListIndex = 1
        cboTextRenderingHint.SetAutomaticRedraws True
        
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
        
    End If
    
    Tools.SetToolBusyState False
    
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
    PDImages.GetActiveImage.GetActiveLayer.SetLayerType PDL_TextBasic
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
    
    'Hide the warning panel and redraw both the viewport, and the UI (as new UI options may now be available)
    Me.UpdateAgainstCurrentLayer
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    Interface.SyncInterfaceToCurrentImage
    
End Sub

Private Sub pdcMain_SizeChanged(Index As Integer)

    'The "text" panel auto-resizes the text entry area to match the size of the container
    If (Index = 0) Then
        txtTextTool.SetSize (pdcMain(Index).GetWidth - txtTextTool.GetLeft) - FixDPI(4), txtTextTool.GetHeight
    End If
    
End Sub

Private Sub sldTextFontSize_Change()

    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer font size
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontSize, sldTextFontSize.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub sldTextFontSize_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontSize, sldTextFontSize.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sldTextFontSize_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FontSize, sldTextFontSize.Value
End Sub

Private Sub sltTextClarity_Change()

    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_TextContrast, sltTextClarity.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

Private Sub sltTextClarity_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_TextContrast, sltTextClarity.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sltTextClarity_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_TextContrast, sltTextClarity.Value
End Sub

Private Sub txtTextTool_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_Text, txtTextTool.Text
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
End Sub

Private Sub txtTextTool_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_Text, txtTextTool.Text, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub txtTextTool_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_Text, txtTextTool.Text
End Sub

'Outside functions can forcibly request an update against the current layer.  If the current layer is a
' non-basic text layer, an option will be displayed to convert the layer.
Public Sub UpdateAgainstCurrentLayer()
    
    'Regardless of layer type, resize our containers to match the current window width.
    Dim winSize As winRect
    If (Not g_WindowManager Is Nothing) Then
        
        g_WindowManager.GetClientWinRect Me.hWnd, winSize
        
        Dim i As Long
        For i = pdcMain.lBound To pdcMain.UBound
            pdcMain(i).SetSize (winSize.x2 - winSize.x1) - pdcMain(i).GetLeft, pdcMain(i).GetHeight
        Next i
        
    End If
    
    If PDImages.IsImageActive() Then

        If PDImages.GetActiveImage.GetActiveLayer.IsLayerText Then
        
            'Check for non-basic-text layers.
            If (PDImages.GetActiveImage.GetActiveLayer.GetLayerType <> PDL_TextBasic) Then
            
                Select Case PDImages.GetActiveImage.GetActiveLayer.GetLayerType()
                
                    Case PDL_TextAdvanced
                        Dim newMessage As String
                        newMessage = g_Language.TranslateMessage("This is an advanced text layer.  To edit it with the basic text tool, you must first convert it to a basic text layer.")
                        newMessage = newMessage & vbCrLf & g_Language.TranslateMessage("(This action is non-destructive.)")
                        Me.lblConvertLayer.Caption = newMessage
                        
                    'In the future, other text layer types can be added here.
                
                End Select
            
                Me.lblConvertLayerConfirm.Caption = g_Language.TranslateMessage("Click here to convert this layer to basic text.")
                
                'Make the prompt panel the size of the tool window
                Me.picConvertLayer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
                
                'Center all labels on the panel
                Me.lblConvertLayer.SetLeft (Me.ScaleWidth - Me.lblConvertLayer.GetWidth) / 2
                Me.lblConvertLayerConfirm.SetLeft (Me.ScaleWidth - Me.lblConvertLayerConfirm.GetWidth) / 2
                
                'Display the panel
                Me.picConvertLayer.Visible = True
                Me.picConvertLayer.ZOrder 0
                
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
    If PDImages.IsImageActive() Then
        If (Not PDImages.GetActiveImage.GetActiveLayer Is Nothing) Then
            CurrentLayerIsText = PDImages.GetActiveImage.GetActiveLayer.IsLayerText
        End If
    End If
    
End Function

Private Sub ChangeMainPanel()
    Dim i As Long
    For i = pdcMain.lBound To pdcMain.UBound
        pdcMain(i).Visible = (i = btsMain.ListIndex)
    Next i
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Update any UI images against the current theme
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(24)
    
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
    Interface.ApplyThemeAndTranslations Me

End Sub
