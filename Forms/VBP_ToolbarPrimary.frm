VERSION 5.00
Begin VB.Form toolbar_Toolbox 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2325
   ClipControls    =   0   'False
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
   ScaleHeight     =   654
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   155
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdLabel lblCategories 
      Height          =   240
      Index           =   0
      Left            =   120
      Top             =   30
      Width           =   2175
      _ExtentX        =   450
      _ExtentY        =   423
      Caption         =   "file"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.jcbutton cmdOpen 
      Height          =   600
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":0000
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Open"
   End
   Begin PhotoDemon.jcbutton cmdSave 
      Height          =   600
      Left            =   60
      TabIndex        =   1
      Top             =   960
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":0D52
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Save"
   End
   Begin PhotoDemon.jcbutton cmdUndo 
      Height          =   600
      Left            =   60
      TabIndex        =   2
      Top             =   1920
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":1AA4
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Undo"
      TooltipBackColor=   -2147483643
   End
   Begin PhotoDemon.jcbutton cmdRedo 
      Height          =   600
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":27F6
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Redo"
      TooltipBackColor=   -2147483643
   End
   Begin PhotoDemon.jcbutton cmdClose 
      Height          =   600
      Left            =   840
      TabIndex        =   4
      Top             =   300
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":3548
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Close"
   End
   Begin PhotoDemon.jcbutton cmdSaveAs 
      Height          =   600
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":429A
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Save as"
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   1
      Left            =   840
      TabIndex        =   9
      Top             =   3000
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   2
      Left            =   1560
      TabIndex        =   10
      Top             =   3000
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   4
      Left            =   840
      TabIndex        =   12
      Top             =   3960
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   5
      Left            =   1560
      TabIndex        =   13
      Top             =   3960
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   7
      Left            =   840
      TabIndex        =   15
      Top             =   4560
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   8
      Left            =   1560
      TabIndex        =   16
      Top             =   4560
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdLabel lblCategories 
      Height          =   240
      Index           =   1
      Left            =   120
      Top             =   1620
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   423
      Caption         =   "undo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.pdLabel lblCategories 
      Height          =   240
      Index           =   2
      Left            =   120
      Top             =   2700
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   423
      Caption         =   "non-destructive"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.pdLabel lblCategories 
      Height          =   240
      Index           =   3
      Left            =   120
      Top             =   3660
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   423
      Caption         =   "selection"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"VBP_ToolbarPrimary.frx":4FEC
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3000
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      UseMnemonic     =   0   'False
      Width           =   2160
      WordWrap        =   -1  'True
   End
   Begin VB.Line lnSeparator 
      BorderColor     =   &H80000002&
      X1              =   8
      X2              =   148
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Label lblRecording 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "macro recording in progress..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   600
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2160
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "toolbar_Toolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Primary Toolbar
'Copyright ©2013-2014 by Tanner Helland
'Created: 02/October/13
'Last updated: 18/October/18
'Last update: start work on an all-new toolbox for the 6.6 release
'
'This form was initially integrated into the main MDI form.  In fall 2014, PhotoDemon left behind the MDI model,
' and all toolbars were moved to their own forms.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub cmdClose_Click()
    Process "Close", True
End Sub

Private Sub cmdOpen_Click()
    Process "Open", True
End Sub

Private Sub cmdRedo_Click()
    Process "Redo", , , UNDO_NOTHING
End Sub

Private Sub cmdSave_Click()
    Process "Save", , , UNDO_NOTHING
End Sub

Private Sub cmdSaveAs_Click()
    Process "Save as", True, , UNDO_NOTHING
End Sub

Private Sub cmdUndo_Click()
    Process "Undo", , , UNDO_NOTHING
End Sub

Private Sub Form_Load()
    
    'Initialize tool button images
    cmdTools(NAV_DRAG).AssignImage "T_HAND"
    cmdTools(NAV_MOVE).AssignImage "T_MOVE"
    cmdTools(QUICK_FIX_LIGHTING).AssignImage "T_NDFX"
    cmdTools(SELECT_RECT).AssignImage "T_SELRECT"
    cmdTools(SELECT_CIRC).AssignImage "T_SELCIRCLE"
    cmdTools(SELECT_LINE).AssignImage "T_SELLINE"
    cmdTools(SELECT_POLYGON).AssignImage "T_SELPOLYGON"
    cmdTools(SELECT_LASSO).AssignImage "T_SELLASSO"
    cmdTools(SELECT_WAND).AssignImage "T_SELWAND"
        
    'Initialize tool button tooltips
    cmdTools(NAV_DRAG).ToolTipText = g_Language.TranslateMessage("Hand (click-and-drag image scrolling)")
    cmdTools(NAV_MOVE).ToolTipText = g_Language.TranslateMessage("Move and resize image layers")
    cmdTools(QUICK_FIX_LIGHTING).ToolTipText = g_Language.TranslateMessage("Apply non-destructive lighting adjustments")
    cmdTools(SELECT_RECT).ToolTipText = g_Language.TranslateMessage("Rectangular Selection")
    cmdTools(SELECT_CIRC).ToolTipText = g_Language.TranslateMessage("Elliptical (Oval) Selection")
    cmdTools(SELECT_LINE).ToolTipText = g_Language.TranslateMessage("Line Selection")
    cmdTools(SELECT_POLYGON).ToolTipText = g_Language.TranslateMessage("Polygon Selection")
    cmdTools(SELECT_LASSO).ToolTipText = g_Language.TranslateMessage("Lasso (Freehand) Selection")
    cmdTools(SELECT_WAND).ToolTipText = g_Language.TranslateMessage("Magic Wand Selection")
        
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.setParentForm Me
    lastUsedSettings.loadAllControlValues
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    lastUsedSettings.saveAllControlValues
    
End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the toggleToolbarVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    
    If g_ProgramShuttingDown Then
        ReleaseFormTheming Me
        g_WindowManager.unregisterForm Me
    Else
        Cancel = True
        toggleToolbarVisibility FILE_TOOLBOX
    End If
    
End Sub

'External functions can use this to re-theme this form at run-time (important when changing languages, for example)
Public Sub requestMakeFormPretty()
    makeFormPretty Me, m_ToolTip
End Sub

'When a new tool is selected, we may need to initialize certain values.
Private Sub newToolSelected()
    
    Select Case g_CurrentTool
    
        'Selection tools
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO
        
            'See if a selection is already active on the image
            If selectionsAllowed(False) Then
            
                'A selection is already active!
                
                'If the existing selection type matches the tool type, no problem - activate the transform tools
                ' (if relevant), but make no other changes to the image
                If (g_CurrentTool = Selection_Handler.getRelevantToolFromSelectShape()) Then
                    metaToggle tSelectionTransform, pdImages(g_CurrentImage).mainSelection.isTransformable
                
                'A selection is already active, and it doesn't match the current tool type!
                Else
                
                    'Handle the special case of circle and rectangular selections, which can be swapped non-destructively.
                    If (g_CurrentTool = SELECT_CIRC) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape = sRectangle) Then
                        pdImages(g_CurrentImage).mainSelection.setSelectionShape sCircle
                        
                    ElseIf (g_CurrentTool = SELECT_RECT) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape = sCircle) Then
                        pdImages(g_CurrentImage).mainSelection.setSelectionShape sRectangle
                        
                    'A selection exists, but it does not match the current tool, and it cannot be non-destructively
                    ' changed to the current type.  Remove it.
                    Else
                        Process "Remove selection", , , UNDO_SELECTION
                    End If
                
                End If
                
            End If
                
        Case Else
        
    End Select
    
    'Finally, because tools may do some custom rendering atop the image canvas, now is a good time to redraw the canvas
    RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub


'External functions can use this to request the selection of a new tool (for example, Select All uses this to set the
' rectangular tool selector as the current tool)
Public Sub selectNewTool(ByVal newToolID As PDTools)
    
    If newToolID <> g_CurrentTool Then
        g_PreviousTool = g_CurrentTool
        g_CurrentTool = newToolID
        resetToolButtonStates
    End If
    
End Sub

'When a new tool button is selected, we need to raise all the others and display the proper options box
Public Sub resetToolButtonStates()
    
    'Start by depressing the selected button and raising all unselected ones
    Dim catID As Long
    For catID = 0 To cmdTools.Count - 1
        If catID = g_CurrentTool Then
            cmdTools(catID).Value = True
        Else
            cmdTools(catID).Value = False
        End If
    Next catID
    
    Dim i As Long
    
    'Next, we need to display the correct tool options panel.  There is no set pattern to this; some tools share
    ' panels, but show/hide certain controls as necessary.  Other tools require their own unique panel.  I've tried
    ' to strike a balance between "as few panels as possible" without going overboard.
    Dim activeToolPanel As Long
    
    Select Case g_CurrentTool
        
        'Move/size tool
        Case NAV_MOVE
            activeToolPanel = 1
        
        'Rectangular, Elliptical, Line selections
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
            activeToolPanel = 3
            
        '"Quick fix" tool(s)
        Case QUICK_FIX_LIGHTING
            activeToolPanel = 2
        
        'If a tool does not require an extra settings panel, set the active panel to -1.  This will hide all panels.
        Case Else
            activeToolPanel = -1
        
    End Select
    
    'If a selection tool is active, we will also need activate a specific subpanel
    Dim activeSelectionSubpanel As Long
    If (getSelectionShapeFromCurrentTool > -1) Then
    
        activeSelectionSubpanel = Selection_Handler.getSelectionSubPanelFromCurrentTool
        
        For i = 0 To toolbar_Options.picSelectionSubcontainer.Count - 1
            If i = activeSelectionSubpanel Then
                toolbar_Options.picSelectionSubcontainer(i).Visible = True
            Else
                toolbar_Options.picSelectionSubcontainer(i).Visible = False
            End If
        Next i
        
    End If
    
    'Check the selection state before swapping tools.  If a selection is active, and the user is switching to the same
    ' tool used to create the current selection, we don't want to erase the current selection.  If they are switching
    ' to a *different* selection tool, however, then we *do* want to erase the current selection.
    If selectionsAllowed(False) And (getRelevantToolFromSelectShape() <> g_CurrentTool) And (getSelectionShapeFromCurrentTool > -1) Then
        
        'Switching between rectangle and circle selections is an exception to the usual rule; these are interchangeable.
        If (g_CurrentTool = SELECT_CIRC) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape = sRectangle) Or _
            (g_CurrentTool = SELECT_RECT) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape = sCircle) Then
            
            'Simply update the shape and redraw the viewport
            pdImages(g_CurrentImage).mainSelection.setSelectionShape Selection_Handler.getSelectionShapeFromCurrentTool
            syncTextToCurrentSelection g_CurrentImage
            RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
            
        Else
            Process "Remove selection", , , UNDO_SELECTION
        End If
        
    End If
    
    'If the current tool is a selection tool, make sure the selection area box (interior/exterior/border) is enabled properly
    If (getSelectionShapeFromCurrentTool > -1) Then toolbar_Options.updateSelectionPanelLayout
    
    'Display the current tool options panel, while hiding all inactive ones.  The On Error Resume statement is used to fix
    ' trouble with the .SetFocus line, below.  That .SetFocus line is helpful for fixing some VB issues with controls embedded
    ' on a picture box (specifically, combo boxes which do not drop-down properly unless a picture box or its child already
    ' has focus).  Sometimes, VB will inexplicably fail to set focus, and it will raise an Error 5 to match; as this is not
    ' a crucial error, just a VB quirk, I don't mind using OERN here.
    On Error Resume Next
    For i = 0 To toolbar_Options.picTools.Count - 1
        If i = activeToolPanel Then
            If Not toolbar_Options.picTools(i).Visible Then
                toolbar_Options.picTools(i).Visible = True
                toolbar_Options.picTools(i).Refresh
                setArrowCursor toolbar_Options.picTools(i)
            End If
            If toolbar_Options.Visible And toolbar_Options.picTools(i).Visible Then toolbar_Options.picTools(i).SetFocus
        Else
            If toolbar_Options.picTools(i).Visible Then toolbar_Options.picTools(i).Visible = False
        End If
    Next i
    
    newToolSelected
        
End Sub

Private Sub cmdTools_Click(Index As Integer)
    
    'Before changing to the new tool, see if the previously active layer has had any non-destructive changes made.
    If Processor.evaluateImageCheckpoint() Then syncInterfaceToCurrentImage
    
    'Update the previous and current tool entries
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = Index
    
    'Update the tool options area to match the newly selected tool
    resetToolButtonStates
    
    'Set a new image checkpoint (necessary to do this manually, as we haven't invoked PD's central processor)
    Processor.setImageCheckpoint
        
End Sub

Private Sub lastUsedSettings_AddCustomPresetData()
    
    'Write the currently selected selection tool to file
    lastUsedSettings.addPresetData "ActiveSelectionTool", g_CurrentTool
    
End Sub

Private Sub lastUsedSettings_ReadCustomPresetData()

    'Restore the last-used selection tool (which will be saved in the main form's preset file, if it exists)
    g_PreviousTool = -1
    If Len(lastUsedSettings.retrievePresetData("ActiveSelectionTool")) > 0 Then
        g_CurrentTool = CLng(lastUsedSettings.retrievePresetData("ActiveSelectionTool"))
    Else
        g_CurrentTool = NAV_DRAG
    End If
    resetToolButtonStates
    
End Sub

