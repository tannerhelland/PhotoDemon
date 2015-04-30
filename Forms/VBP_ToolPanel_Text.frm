VERSION 5.00
Begin VB.Form toolpanel_Text 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
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
   ScaleWidth      =   1110
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.colorSelector csTextFontColor 
      Height          =   420
      Left            =   7680
      TabIndex        =   1
      Top             =   930
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   741
      curColor        =   0
   End
   Begin PhotoDemon.textUpDown tudTextFontSize 
      Height          =   345
      Left            =   7680
      TabIndex        =   2
      Top             =   510
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
      Min             =   1
      Max             =   1000
      Value           =   16
   End
   Begin PhotoDemon.pdTextBox txtTextTool 
      Height          =   1380
      Left            =   840
      TabIndex        =   3
      Top             =   30
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2434
      FontSize        =   9
      Multiline       =   -1  'True
   End
   Begin PhotoDemon.pdComboBox cboTextFontFace 
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   60
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   0
      Left            =   12600
      Top             =   1080
      Width           =   2445
      _ExtentX        =   0
      _ExtentY        =   503
      Caption         =   "(this tool is under construction)"
      ForeColor       =   255
      UseCustomForeColor=   -1  'True
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
      Caption         =   "font color:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdComboBox cboTextRenderingHint 
      Height          =   375
      Left            =   11760
      TabIndex        =   5
      Top             =   60
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   5
      Left            =   10200
      Top             =   120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "AA and hinting:"
      ForeColor       =   0
   End
   Begin PhotoDemon.textUpDown tudTextClarity 
      Height          =   345
      Left            =   11760
      TabIndex        =   0
      Top             =   510
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
      Max             =   12
      Value           =   4
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   6
      Left            =   10200
      Top             =   570
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "AA clarity:"
      ForeColor       =   0
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

Private Sub tudTextClarity_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_TextContrast, tudTextClarity.Value
    
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

Private Sub txtTextTool_GotFocus()
    'Disable accelerators
    FormMain.ctlAccelerator.Enabled = False
    Debug.Print "lost focus"
End Sub

Private Sub txtTextTool_LostFocus()
    'Re-enable accelerators
    FormMain.ctlAccelerator.Enabled = True
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
