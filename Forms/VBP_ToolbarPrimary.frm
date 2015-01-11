VERSION 5.00
Begin VB.Form toolbar_Toolbox 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2085
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
   ScaleWidth      =   139
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdLabel lblCategories 
      Height          =   240
      Index           =   0
      Left            =   120
      Top             =   30
      Width           =   2175
      _ExtentX        =   450
      _ExtentY        =   503
      Caption         =   "file"
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   2880
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   2880
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   4
      Left            =   840
      TabIndex        =   5
      Top             =   3840
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   5
      Left            =   1560
      TabIndex        =   6
      Top             =   3840
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   7
      Left            =   840
      TabIndex        =   8
      Top             =   4440
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   8
      Left            =   1560
      TabIndex        =   9
      Top             =   4440
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
      _ExtentY        =   503
      Caption         =   "undo"
   End
   Begin PhotoDemon.pdLabel lblCategories 
      Height          =   240
      Index           =   2
      Left            =   120
      Top             =   2580
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      Caption         =   "non-destructive"
   End
   Begin PhotoDemon.pdLabel lblCategories 
      Height          =   240
      Index           =   3
      Left            =   120
      Top             =   3540
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      Caption         =   "selection"
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   300
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   300
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   2
      Left            =   1560
      TabIndex        =   12
      Top             =   300
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   4
      Left            =   840
      TabIndex        =   14
      Top             =   960
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   5
      Left            =   1560
      TabIndex        =   15
      Top             =   960
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   7
      Left            =   840
      TabIndex        =   17
      Top             =   1920
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   8
      Left            =   1560
      TabIndex        =   18
      Top             =   1920
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
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
      TabIndex        =   0
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
'Copyright 2013-2015 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 18/Oct/14
'Last update: start work on an all-new toolbox for the 6.6 release
'
'This form was initially integrated into the main MDI form.  In fall 2013, PhotoDemon left behind the MDI model,
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

'When we are responsible for this window resizing (because the user is resizing our window manually), we set this to TRUE.
' This variable is then checked before requesting additional redraws during our resize event.
Private m_WeAreResponsibleForResize As Boolean

'The user has various options for how this toolbox is displayed.  When changed, those options need to update these
' module-level values, then manually reflow the interface.
Private m_ShowCategoryLabels As Boolean

'Button size is also controllable.  In the future, this will result in an actual change to the images used on the buttons.
' For now, however, we simply resize the buttons themselves.
Private Enum PD_TOOLBOX_BUTTON_SIZE
    PDTBS_SMALL = 0
    PDTBS_MEDIUM = 1
    PDTBS_LARGE = 2
End Enum

Private m_ButtonSize As PD_TOOLBOX_BUTTON_SIZE

Private Const BTS_WIDTH_SMALL As Long = 32
Private Const BTS_WIDTH_MEDIUM As Long = 46
Private Const BTS_WIDTH_LARGE As Long = 56

Private Const BTS_HEIGHT_SMALL As Long = 32
Private Const BTS_HEIGHT_MEDIUM As Long = 38
Private Const BTS_HEIGHT_LARGE As Long = 48

Private m_ButtonWidth As Long, m_ButtonHeight As Long

Private Sub cmdFile_Click(Index As Integer)
    
    Select Case Index
    
        Case FILE_NEW
            Process "New image", True
        
        Case FILE_OPEN
            Process "Open", True
        
        Case FILE_CLOSE
            Process "Close", True
        
        Case FILE_SAVE
            Process "Save", , , UNDO_NOTHING
        
        Case FILE_SAVEAS_LAYERS
            Process "Save copy", , , UNDO_NOTHING
            
        Case FILE_SAVEAS_FLAT
            Process "Save as", True, , UNDO_NOTHING
        
        Case FILE_UNDO
            Process "Undo", , , UNDO_NOTHING
        
        Case FILE_FADE
            Process "Fade", True
        
        Case FILE_REDO
            Process "Redo", , , UNDO_NOTHING
    
    End Select
    
    cmdFile(Index).notifyFocusLost
    cmdFile(Index).Value = False

End Sub

Private Sub Form_Load()
    
    'Initialize file tool button images
    cmdFile(FILE_NEW).AssignImage "TF_NEW"
    cmdFile(FILE_OPEN).AssignImage "TF_OPEN", , , 10
    cmdFile(FILE_CLOSE).AssignImage "TF_CLOSE", , 100
    cmdFile(FILE_SAVE).AssignImage "TF_SAVE", , 50
    cmdFile(FILE_SAVEAS_LAYERS).AssignImage "TF_SAVEPDI", , 50
    cmdFile(FILE_SAVEAS_FLAT).AssignImage "TF_SAVEAS", , 50
    cmdFile(FILE_UNDO).AssignImage "TF_UNDO", , 50
    cmdFile(FILE_FADE).AssignImage "TF_FADE", , 50
    cmdFile(FILE_REDO).AssignImage "TF_REDO"
            
    'Initialize canvas tool button images
    cmdTools(NAV_DRAG).AssignImage "T_HAND"
    cmdTools(NAV_MOVE).AssignImage "T_MOVE"
    cmdTools(QUICK_FIX_LIGHTING).AssignImage "T_NDFX"
    cmdTools(SELECT_RECT).AssignImage "T_SELRECT"
    cmdTools(SELECT_CIRC).AssignImage "T_SELCIRCLE"
    cmdTools(SELECT_LINE).AssignImage "T_SELLINE"
    cmdTools(SELECT_POLYGON).AssignImage "T_SELPOLYGON"
    cmdTools(SELECT_LASSO).AssignImage "T_SELLASSO"
    cmdTools(SELECT_WAND).AssignImage "T_SELWAND"
        
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.setParentForm Me
    lastUsedSettings.loadAllControlValues
        
    'Retrieve any relevant toolbox display settings from the user's preferences file
    m_ShowCategoryLabels = g_UserPreferences.GetPref_Boolean("Core", "Show Toolbox Category Labels", True)
    m_ButtonSize = g_UserPreferences.GetPref_Long("Core", "Toolbox Button Size", 1)
    
    'Note that we don't actually reflow the interface here; that will happen later, when the form's previous size and
    ' position are loaded from the user's preference file.
    
    'As a final step, redraw everything against the current theme.
    updateAgainstCurrentTheme
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If the mouse is near the resizable edge of the toolbar (the left edge, currently), allow the user to resize
    ' the layer toolbox.
    Dim mouseInResizeTerritory As Boolean
    
    'How close does the mouse have to be to the form border to allow resizing; currently we use 7 pixels, while accounting
    ' for DPI variance (e.g. 7 pixels at 96 dpi)
    Dim resizeBorderAllowance As Long
    resizeBorderAllowance = fixDPI(7)
    
    Dim hitCode As Long
    
    'Check the mouse position to see if it's in resize territory (along the left edge of the toolbox)
    If (y > 0) And (y < Me.ScaleHeight) And (x > (Me.ScaleWidth - resizeBorderAllowance)) Then
        mouseInResizeTerritory = True
        hitCode = HTRIGHT
    End If
    
    'If the left mouse button is down, and the mouse is in resize territory, initiate an API resize event
    If mouseInResizeTerritory Then
    
        'Change the cursor to a resize cursor
        setSizeWECursor Me
        
        If (Button = vbLeftButton) Then
            m_WeAreResponsibleForResize = True
            ReleaseCapture
            SendMessage Me.hWnd, WM_NCLBUTTONDOWN, hitCode, ByVal 0&
            
            'A premature exit is required, because the end of this sub contains code to detect the release of the
            ' mouse after a drag event.  Because the event is not being initiated normally, we can't detect a standard
            ' MouseUp event, so instead, we mimic it by checking MouseMove and m_WeAreResponsibleForResize = TRUE.
            Exit Sub
            
        End If
        
    Else
        setArrowCursor Me
    End If
    
    'Check for mouse release; we will only reach this point if the mouse is *not* in resize territory, which in turn
    ' means we can free the release code and resize the window now.  (On some OS/theme combinations, the canvas will
    ' live-resize as the mouse is moved.  On others, the canvas won't redraw until the mouse is released.)
    If m_WeAreResponsibleForResize Then
        
        m_WeAreResponsibleForResize = False
        
        'If theming is disabled, window performance is so poor that the window manager will automatically
        ' disable canvas updates until the mouse is released.  Request a full update now.
        If (Not g_IsThemingEnabled) Then g_WindowManager.notifyToolboxResized Me.hWnd, True
        
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    lastUsedSettings.saveAllControlValues
    
End Sub

Private Sub Form_Resize()
    
    'Reflow the form's contents
    reflowToolboxLayout
    
End Sub

Private Sub reflowToolboxLayout()

    Dim i As Long
    
    'Reflow label width first; they are easy because they simply match the width of the form
    For i = 0 To lblCategories.UBound
        lblCategories(i).Width = Me.ScaleWidth - (lblCategories(i).Left + fixDPI(2))
    Next i
    
    lblRecording.Width = Me.ScaleWidth - (lblRecording.Left + fixDPI(2))
    
    'Next, we are going to reflow the interface in two segments: the "file" buttons (which are handled separately, since
    ' they are actual buttons and not persistent toggles), then the toolbox buttons.
    
    'Conceptually, reflowing is simple: we iterate through controls in top-to-bottom order, and we position them
    ' according to a few simple rules:
    ' 1) Title labels are handled first.  They always receive their own row
    ' 2) Buttons are laid out in groups.  The groups are hand-coded.
    ' 3) Buttons are laid out in horizontal rows until the end of the form is reached.  When this happens, buttons are
    '     pushed down to a new row.
    ' 4) We repeat the pattern until all buttons and labels have been dealt with.
    
    Dim hOffset As Long, vOffset As Long
    Dim hOffsetDefaultLabel As Long, hOffsetDefaultButton As Long
    Dim labelMarginBottom As Long, labelMarginTop As Long, buttonMarginBottom As Long, buttonMarginRight As Long
    
    'Start by establishing default values
    hOffsetDefaultLabel = fixDPI(4) 'Left-position of labels
    hOffsetDefaultButton = 0        'Left-position of left-most buttons
    labelMarginBottom = fixDPI(4)   'Distance between the bottom of labels and the top of buttons
    labelMarginTop = fixDPI(2)      'Distance between the bottom of buttons and top of labels
    buttonMarginBottom = 0          'Distance between button rows
    buttonMarginRight = 0           'Distance between buttons
    
    'If category labels are displayed, make them visible now
    For i = 0 To lblCategories.UBound
        If lblCategories(i).Visible <> m_ShowCategoryLabels Then lblCategories(i).Visible = m_ShowCategoryLabels
    Next i
    
    'File group
    If m_ShowCategoryLabels Then
        lblCategories(0).Move hOffsetDefaultLabel, fixDPI(2)
        vOffset = lblCategories(0).Top + lblCategories(0).Height + labelMarginBottom
    Else
        vOffset = fixDPI(2)
    End If
    
    For i = FILE_NEW To FILE_SAVEAS_FLAT
        
        'Move this button into position
        cmdFile(i).Move hOffset, vOffset
        
        'Calculate a new offset for the next button
        hOffset = hOffset + cmdFile(i).Width + buttonMarginRight
        If hOffset + cmdFile(i).Width > Me.ScaleWidth Then
            hOffset = hOffsetDefaultButton
            vOffset = vOffset + cmdFile(i).Height + buttonMarginBottom
        End If
        
    Next i
    
    'Undo group
    If m_ShowCategoryLabels Then
    
        If vOffset < cmdFile(FILE_SAVEAS_FLAT).Top + cmdFile(FILE_SAVEAS_FLAT).Height Then
            vOffset = cmdFile(FILE_SAVEAS_FLAT).Top + cmdFile(FILE_SAVEAS_FLAT).Height + buttonMarginBottom
        End If
        
        vOffset = vOffset + labelMarginTop
        lblCategories(1).Move hOffsetDefaultLabel, vOffset
        vOffset = lblCategories(1).Top + lblCategories(1).Height + labelMarginBottom
        hOffset = hOffsetDefaultButton
        
    End If
    
    For i = FILE_UNDO To FILE_REDO
        
        'Move this button into position
        cmdFile(i).Move hOffset, vOffset
        
        'Calculate a new offset for the next button
        hOffset = hOffset + cmdFile(i).Width + buttonMarginRight
        If hOffset + cmdFile(i).Width > Me.ScaleWidth Then
            hOffset = hOffsetDefaultButton
            vOffset = vOffset + cmdFile(i).Height + buttonMarginBottom
        End If
        
    Next i
    
    'Non-destructive group
    If m_ShowCategoryLabels Then
    
        If vOffset < cmdFile(FILE_REDO).Top + cmdFile(FILE_REDO).Height Then
            vOffset = cmdFile(FILE_REDO).Top + cmdFile(FILE_REDO).Height + buttonMarginBottom
        End If
        
        vOffset = vOffset + labelMarginTop
        lblCategories(2).Move hOffsetDefaultLabel, vOffset
        vOffset = lblCategories(2).Top + lblCategories(2).Height + labelMarginBottom
        hOffset = hOffsetDefaultButton
        
    End If
    
    For i = NAV_DRAG To QUICK_FIX_LIGHTING
        
        'Move this button into position
        cmdTools(i).Move hOffset, vOffset
        
        'Calculate a new offset for the next button
        hOffset = hOffset + cmdTools(i).Width + buttonMarginRight
        If hOffset + cmdTools(i).Width > Me.ScaleWidth Then
            hOffset = hOffsetDefaultButton
            vOffset = vOffset + cmdTools(i).Height + buttonMarginBottom
        End If
        
    Next i
    
    'Selection group
    If m_ShowCategoryLabels Then
    
        If vOffset < cmdTools(QUICK_FIX_LIGHTING).Top + cmdTools(QUICK_FIX_LIGHTING).Height Then
            vOffset = cmdTools(QUICK_FIX_LIGHTING).Top + cmdTools(QUICK_FIX_LIGHTING).Height + buttonMarginBottom
        End If
        
        vOffset = vOffset + labelMarginTop
        lblCategories(3).Move hOffsetDefaultLabel, vOffset
        vOffset = lblCategories(3).Top + lblCategories(3).Height + labelMarginBottom
        hOffset = hOffsetDefaultButton
        
    End If
    
    For i = SELECT_RECT To SELECT_WAND
        
        'Move this button into position
        cmdTools(i).Move hOffset, vOffset
        
        'Calculate a new offset for the next button
        hOffset = hOffset + cmdTools(i).Width + buttonMarginRight
        If hOffset + cmdTools(i).Width > Me.ScaleWidth Then
            hOffset = hOffsetDefaultButton
            vOffset = vOffset + cmdTools(i).Height + buttonMarginBottom
        End If
        
    Next i
    
    'Macro recording message
    If vOffset < cmdTools(SELECT_WAND).Top + cmdTools(SELECT_WAND).Height Then
        vOffset = cmdTools(SELECT_WAND).Top + cmdTools(SELECT_WAND).Height + buttonMarginBottom
    End If
    
    vOffset = vOffset + labelMarginTop
    lblRecording.Move lblRecording.Left, vOffset

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
    makeFormPretty Me
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
    Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
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
    
    'Next, we can automatically hide the options toolbox for certain tools (because they have no options).  This is a
    ' nice courtesy, as it frees up space on the main canvas area if the current tool has no adjustable options.
    If Me.Visible Then
    
        Select Case g_CurrentTool
            
            'Hand tool is currently the only tool without additional options
            Case NAV_DRAG
                g_WindowManager.setWindowVisibility toolbar_Options.hWnd, False, False
                
            'All other tools expose options, so display the toolbox (unless the user has disabled the window completely)
            Case Else
                g_WindowManager.setWindowVisibility toolbar_Options.hWnd, g_UserPreferences.GetPref_Boolean("Core", "Show Selections Toolbox", True), False
                
        End Select
        
    End If
    
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
            Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
            
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
    
    If Len(lastUsedSettings.retrievePresetData("ActiveSelectionTool")) <> 0 Then
        g_CurrentTool = CLng(lastUsedSettings.retrievePresetData("ActiveSelectionTool"))
    Else
        g_CurrentTool = NAV_DRAG
    End If
    
    resetToolButtonStates
    
End Sub

'Used to change the visibility of category labels.  When disabled, the button layout is reflowed into a continuous
' stream of buttons.  When enabled, buttons are sorted by category.
Public Sub toggleToolCategoryLabels()
    
    FormMain.MnuWindowToolbox(2).Checked = Not FormMain.MnuWindowToolbox(2).Checked
    g_UserPreferences.SetPref_Boolean "Core", "Show Toolbox Category Labels", CBool(FormMain.MnuWindowToolbox(2).Checked)
    m_ShowCategoryLabels = CBool(FormMain.MnuWindowToolbox(2).Checked)
    
    'Reflow the interface
    reflowToolboxLayout
    
End Sub

'Used to change the display size of toolbox buttons.  newSize is expected on the range [0, 2] for small, medium, large
Public Sub updateButtonSize(ByVal newSize As Long, Optional ByVal suppressRedraw As Boolean = False)
    
    'Export the updated size to file
    If Not suppressRedraw Then g_UserPreferences.SetPref_Long "Core", "Toolbox Button Size", newSize
    
    'Update our internal size metrics to match
    m_ButtonSize = newSize
    
    Select Case m_ButtonSize
    
        Case PDTBS_SMALL
            m_ButtonWidth = fixDPI(BTS_WIDTH_SMALL)
            m_ButtonHeight = fixDPI(BTS_HEIGHT_SMALL)
        
        Case PDTBS_MEDIUM
            m_ButtonWidth = fixDPI(BTS_WIDTH_MEDIUM)
            m_ButtonHeight = fixDPI(BTS_HEIGHT_MEDIUM)
        
        Case PDTBS_LARGE
            m_ButtonWidth = fixDPI(BTS_WIDTH_LARGE)
            m_ButtonHeight = fixDPI(BTS_HEIGHT_LARGE)
    
    End Select
    
    'Update all buttons to match
    Dim i As Long
    
    For i = 0 To cmdFile.UBound
        If (cmdFile(i).Width <> m_ButtonWidth) Or (cmdFile(i).Height <> m_ButtonHeight) Then
            cmdFile(i).Move cmdFile(i).Left, cmdFile(i).Top, m_ButtonWidth, m_ButtonHeight
        End If
    Next i
    
    For i = 0 To cmdTools.UBound
        If (cmdTools(i).Width <> m_ButtonWidth) Or (cmdTools(i).Height <> m_ButtonHeight) Then
            cmdTools(i).Move cmdTools(i).Left, cmdTools(i).Top, m_ButtonWidth, m_ButtonHeight
        End If
    Next i
    
    'Convert the newSize value to a menu index
    newSize = newSize + 4

    'Apply a checkbox to the matching menu item
    For i = 4 To 6
        If i = newSize Then
            FormMain.MnuWindowToolbox(i).Checked = True
        Else
            FormMain.MnuWindowToolbox(i).Checked = False
        End If
    Next i
    
    g_WindowManager.updateMinimumDimensions Me.hWnd, m_ButtonWidth
    
    'Reflow the interface as requested
    If Not suppressRedraw Then reflowToolboxLayout
    
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
    
    'Tooltips must be manually re-assigned according to the current language.  This is a necessary evil, if the user switches
    ' between two non-English languages at run-time.
    
    'File tool buttons come first
    cmdFile(FILE_NEW).assignTooltip g_Language.TranslateMessage("This option will create a blank image.  Other ways to create new images can be found in the File -> Import menu."), g_Language.TranslateMessage("New Image")
    cmdFile(FILE_OPEN).assignTooltip g_Language.TranslateMessage("Another way to open images is dragging them from your desktop or Windows Explorer and dropping them onto PhotoDemon."), g_Language.TranslateMessage("Open one or more images for editing")
    
    If g_ConfirmClosingUnsaved Then
        cmdFile(FILE_CLOSE).assignTooltip g_Language.TranslateMessage("If the current image has not been saved, you will receive a prompt to save it before it closes."), g_Language.TranslateMessage("Close the current image")
    Else
        cmdFile(FILE_CLOSE).assignTooltip g_Language.TranslateMessage("Because you have turned off save prompts (via Edit -> Preferences), you WILL NOT receive a prompt to save this image before it closes."), g_Language.TranslateMessage("Close the current image")
    End If
    
    If g_UserPreferences.GetPref_Long("Saving", "Overwrite Or Copy", 0) = 0 Then
        cmdFile(FILE_SAVE).assignTooltip g_Language.TranslateMessage("WARNING: this will overwrite the current image file.  To save to a different file, use the ""Save As"" button."), g_Language.TranslateMessage("Save image in current format")
    Else
        cmdFile(FILE_SAVE).assignTooltip g_Language.TranslateMessage("You have specified ""safe"" save mode, which means that each save will create a new file with an auto-incremented filename."), g_Language.TranslateMessage("Save image in current format")
    End If
    
    cmdFile(FILE_SAVEAS_LAYERS).assignTooltip g_Language.TranslateMessage("Use this to quickly save a lossless copy of the current image.  The lossless copy will be saved in PDI format, in the image's current folder, using the current filename (plus an auto-incremented number, as necessary)."), g_Language.TranslateMessage("Save lossless copy")
    cmdFile(FILE_SAVEAS_FLAT).assignTooltip g_Language.TranslateMessage("The Save As command always raises a dialog, so you can specify a new file name, folder, and/or image format for the current image."), g_Language.TranslateMessage("Save As (export to new format or filename)")
    
    cmdFile(FILE_UNDO).assignTooltip g_Language.TranslateMessage("Undo last action")
    cmdFile(FILE_FADE).assignTooltip g_Language.TranslateMessage("Fade last action")
    cmdFile(FILE_REDO).assignTooltip g_Language.TranslateMessage("Redo previous action")
    
    'Painting tool buttons are next
    cmdTools(NAV_DRAG).assignTooltip g_Language.TranslateMessage("Hand (click-and-drag image scrolling)")
    cmdTools(NAV_MOVE).assignTooltip g_Language.TranslateMessage("Move and resize image layers")
    cmdTools(QUICK_FIX_LIGHTING).assignTooltip g_Language.TranslateMessage("Apply non-destructive lighting adjustments")
    cmdTools(SELECT_RECT).assignTooltip g_Language.TranslateMessage("Rectangular Selection")
    cmdTools(SELECT_CIRC).assignTooltip g_Language.TranslateMessage("Elliptical (Oval) Selection")
    cmdTools(SELECT_LINE).assignTooltip g_Language.TranslateMessage("Line Selection")
    cmdTools(SELECT_POLYGON).assignTooltip g_Language.TranslateMessage("Polygon Selection")
    cmdTools(SELECT_LASSO).assignTooltip g_Language.TranslateMessage("Lasso (Freehand) Selection")
    cmdTools(SELECT_WAND).assignTooltip g_Language.TranslateMessage("Magic Wand Selection")
    
End Sub
