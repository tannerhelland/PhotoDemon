VERSION 5.00
Begin VB.Form toolbar_Toolbox 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2115
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
   ScaleWidth      =   141
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
      AutoToggle      =   -1  'True
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
      AutoToggle      =   -1  'True
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
      AutoToggle      =   -1  'True
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
      AutoToggle      =   -1  'True
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
      AutoToggle      =   -1  'True
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
      AutoToggle      =   -1  'True
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
      AutoToggle      =   -1  'True
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
      AutoToggle      =   -1  'True
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
      AutoToggle      =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblCategories 
      Height          =   240
      Index           =   4
      Left            =   120
      Top             =   5100
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      Caption         =   "text"
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   10
      Left            =   840
      TabIndex        =   20
      Top             =   5400
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdLabel lblCategories 
      Height          =   240
      Index           =   5
      Left            =   120
      Top             =   6120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      Caption         =   "paint"
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   11
      Left            =   120
      TabIndex        =   21
      Top             =   6480
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
   End
   Begin VB.Line lnRightSeparator 
      X1              =   136
      X2              =   136
      Y1              =   0
      Y2              =   648
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
      Top             =   7200
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

'These values are basically constants; they are set by the ReflowToolboxLayout function.
Private m_hOffsetDefaultLabel As Long, m_hOffsetDefaultButton As Long
Private m_labelMarginBottom As Long, m_labelMarginTop As Long
Private m_buttonMarginBottom As Long, m_buttonMarginRight As Long
Private m_rightBoundary As Long

'The currently active tool panel and ID key will be mirrored to this reference
Private m_ActiveToolPanelKey As String

'This form supports a variety of resize modes, and we use cMouseEvents to handle cursor duties
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

Private Sub cmdFile_Click(Index As Integer)
        
    'If the user is dragging the mouse in from the right, and the toolbox has been shrunk from its default setting, the class cursor
    ' for forms may get stuck on the west/east "resize" cursor.  To avoid this, reset it after any button click.
    cMouseEvents.setSystemCursor IDC_ARROW
        
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
    
End Sub

'When the mouse leaves this toolbox, reset it to an arrow (so other forms don't magically acquire the west/east resize cursor, as the mouse is
' likely to leave off the right side of this form)
Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    cMouseEvents.setSystemCursor IDC_ARROW
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'If the mouse is near the resizable edge of the toolbar (the left edge, currently), allow the user to resize
    ' the layer toolbox.
    Dim mouseInResizeTerritory As Boolean
    
    'How close does the mouse have to be to the form border to allow resizing; currently we use 7 pixels, while accounting
    ' for DPI variance (e.g. 7 pixels at 96 dpi)
    Dim resizeBorderAllowance As Long
    resizeBorderAllowance = FixDPI(7)
    
    Dim hitCode As Long
    
    'Check the mouse position to see if it's in resize territory (along the left edge of the toolbox)
    If (y > 0) And (y < Me.ScaleHeight) And (x > (Me.ScaleWidth - resizeBorderAllowance)) Then
        mouseInResizeTerritory = True
        hitCode = HTRIGHT
    End If
    
    'If the left mouse button is down, and the mouse is in resize territory, initiate an API resize event
    If mouseInResizeTerritory Then
    
        'Change the cursor to a resize cursor
        cMouseEvents.setSystemCursor IDC_SIZEWE
        
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
        cMouseEvents.setSystemCursor IDC_ARROW
    End If
    
    'Check for mouse release; we will only reach this point if the mouse is *not* in resize territory, which in turn
    ' means we can free the release code and resize the window now.  (On some OS/theme combinations, the canvas will
    ' live-resize as the mouse is moved.  On others, the canvas won't redraw until the mouse is released.)
    If m_WeAreResponsibleForResize Then
        
        m_WeAreResponsibleForResize = False
        
        'If theming is disabled, window performance is so poor that the window manager will automatically
        ' disable canvas updates until the mouse is released.  Request a full update now.
        If (Not g_IsThemingEnabled) Then g_WindowManager.NotifyToolboxResized Me.hWnd, True
        
    End If

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
    
    cmdTools(VECTOR_TEXT).AssignImage "TV_TEXT", , , 50
    cmdTools(VECTOR_FANCYTEXT).AssignImage "TV_FANCYTEXT", , , 50
    
    cmdTools(PAINT_BASICBRUSH).AssignImage "PNT_BASICBRUSH"
    
    'Initialize a mouse handler
    Set cMouseEvents = New pdInputMouse
    cMouseEvents.addInputTracker Me.hWnd, True, True, , True
    
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
    UpdateAgainstCurrentTheme
    
End Sub

Private Sub Form_LostFocus()
    cMouseEvents.setSystemCursor IDC_ARROW
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    lastUsedSettings.saveAllControlValues
    
End Sub

Private Sub Form_Resize()
    
    'Reflow the form's contents
    ReflowToolboxLayout
    
End Sub

Private Sub ReflowToolboxLayout()

    Dim i As Long
    
    'Before doing anything complicated, right-align the line separator
    lnRightSeparator.x1 = Me.ScaleWidth - 1
    lnRightSeparator.y1 = 0
    lnRightSeparator.x2 = lnRightSeparator.x1
    lnRightSeparator.y2 = Me.ScaleHeight
    
    'We're also going to mark the right boundary for images, which allows for some padding when reflowing the interface
    m_rightBoundary = Me.ScaleWidth - 3
    
    'Reflow label width first; they are easy because they simply match the width of the form
    For i = 0 To lblCategories.UBound
        lblCategories(i).Width = m_rightBoundary - (lblCategories(i).Left + FixDPI(2))
    Next i
    
    lblRecording.Width = m_rightBoundary - (lblRecording.Left + FixDPI(2))
    
    'Next, we are going to reflow the interface in two segments: the "file" buttons (which are handled separately, since
    ' they are actual buttons and not persistent toggles), then the toolbox buttons.
    
    'Conceptually, reflowing is simple: we iterate through controls in top-to-bottom order, and we position them
    ' according to a few simple rules:
    ' 1) Title labels are handled first.  They always receive their own row.
    ' 2) Buttons are laid out in groups.  The groups are hand-coded.
    ' 3) Buttons are laid out in horizontal rows until the end of the form is reached.  When this happens, buttons are
    '     pushed down to a new row.
    ' 4) We repeat the pattern until all buttons and labels have been dealt with.
    
    Dim hOffset As Long, vOffset As Long
    
    'Start by establishing default values
    m_hOffsetDefaultLabel = FixDPI(4) 'Left-position of labels
    m_hOffsetDefaultButton = 0        'Left-position of left-most buttons
    m_labelMarginBottom = FixDPI(4)   'Distance between the bottom of labels and the top of buttons
    m_labelMarginTop = FixDPI(2)      'Distance between the bottom of buttons and top of labels
    m_buttonMarginBottom = 0          'Distance between button rows
    m_buttonMarginRight = 0           'Distance between buttons
    
    'If category labels are displayed, make them visible now
    For i = 0 To lblCategories.UBound
        If lblCategories(i).Visible <> m_ShowCategoryLabels Then lblCategories(i).Visible = m_ShowCategoryLabels
    Next i
    
    'File group.  We position this label manually, and it serves as the reference for all subsequent labels.
    If m_ShowCategoryLabels Then
        lblCategories(0).Move m_hOffsetDefaultLabel, FixDPI(2)
        vOffset = lblCategories(0).Top + lblCategories(0).Height + m_labelMarginBottom
    Else
        vOffset = FixDPI(2)
    End If
    
    ReflowButtonSet False, FILE_NEW, FILE_SAVEAS_FLAT, hOffset, vOffset
    
    'Undo group
    PositionToolLabel lblCategories(1), cmdFile(FILE_SAVEAS_FLAT), hOffset, vOffset
    ReflowButtonSet False, FILE_UNDO, FILE_REDO, hOffset, vOffset
    
    'Non-destructive group
    PositionToolLabel lblCategories(2), cmdFile(FILE_REDO), hOffset, vOffset
    ReflowButtonSet True, NAV_DRAG, QUICK_FIX_LIGHTING, hOffset, vOffset
    
    'Selection group
    PositionToolLabel lblCategories(3), cmdTools(QUICK_FIX_LIGHTING), hOffset, vOffset
    ReflowButtonSet True, SELECT_RECT, SELECT_WAND, hOffset, vOffset
    
    'Vector group
    PositionToolLabel lblCategories(4), cmdTools(SELECT_WAND), hOffset, vOffset
    ReflowButtonSet True, VECTOR_TEXT, VECTOR_FANCYTEXT, hOffset, vOffset
    
    'Paint group
    PositionToolLabel lblCategories(5), cmdTools(VECTOR_FANCYTEXT), hOffset, vOffset
    ReflowButtonSet True, PAINT_BASICBRUSH, PAINT_BASICBRUSH, hOffset, vOffset
        
    'Macro recording message
    If vOffset < cmdTools(cmdTools.UBound).Top + cmdTools(cmdTools.UBound).Height Then
        vOffset = cmdTools(cmdTools.UBound).Top + cmdTools(cmdTools.UBound).Height + m_buttonMarginBottom
    End If
    
    vOffset = vOffset + m_labelMarginTop
    lblRecording.Move lblRecording.Left, vOffset

End Sub

'Companion function to ReflowToolboxLayout(), above.
Private Sub PositionToolLabel(ByRef targetLabel As Object, ByRef referenceButton As Object, ByRef hOffset As Long, ByRef vOffset As Long)
    
    If m_ShowCategoryLabels Then
    
        If vOffset < referenceButton.Top + referenceButton.Height Then
            vOffset = referenceButton.Top + referenceButton.Height + m_buttonMarginBottom
        End If
        
        vOffset = vOffset + m_labelMarginTop
        targetLabel.Move m_hOffsetDefaultLabel, vOffset
        vOffset = targetLabel.Top + targetLabel.Height + m_labelMarginBottom
        hOffset = m_hOffsetDefaultButton
        
    End If
    
End Sub

'Companion function to ReflowToolboxLayout(), above.
Private Sub ReflowButtonSet(ByVal categoryIsTools As Boolean, ByVal startIndex As Long, ByVal endIndex As Long, ByRef hOffset As Long, ByRef vOffset As Long)
    
    Dim i As Long, targetObject As Object
        
    For i = startIndex To endIndex
        
        'Toolbox buttons are divided into two groups: file, and canvas.  The code for positioning these is identical,
        ' so changes to one block should be mirrored to the other.
        If categoryIsTools Then
            Set targetObject = cmdTools(i)
        Else
            Set targetObject = cmdFile(i)
        End If
        
        'Move this button into position
        targetObject.Move hOffset, vOffset
        
        'Calculate a new offset for the next button
        hOffset = hOffset + targetObject.Width + m_buttonMarginRight
        If hOffset + targetObject.Width > m_rightBoundary Then
            hOffset = m_hOffsetDefaultButton
            vOffset = vOffset + targetObject.Height + m_buttonMarginBottom
        End If
        
    Next i
    
End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the toggleToolbarVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    
    If g_ProgramShuttingDown Then
        ReleaseFormTheming Me
        g_WindowManager.UnregisterForm Me
    Else
        Cancel = True
        ToggleToolbarVisibility FILE_TOOLBOX
    End If
    
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
                    MetaToggle tSelectionTransform, pdImages(g_CurrentImage).mainSelection.isTransformable
                
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
            
            'Finally, because tools may do some custom rendering atop the image canvas, now is a good time to redraw the canvas.
            ' (Note that we can use a very late pipeline stage, as only tool-specific overlays need to be redrawn.)
            Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                
        Case VECTOR_TEXT, VECTOR_FANCYTEXT
        
            'Switching text tools may require us to redraw the text buffer, as the font rendering engine changes depending
            ' on the current text tool.
            Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        Case Else
        
            'Finally, because tools may do some custom rendering atop the image canvas, now is a good time to redraw the canvas.
            ' (Note that we can use a very late pipeline stage, as only tool-specific overlays need to be redrawn.)
            Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
    End Select
        
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
    Dim activeToolPanel As Form
    
    Select Case g_CurrentTool
        
        'Move/size tool
        Case NAV_MOVE
            Load toolpanel_MoveSize
            m_ActiveToolPanelKey = "MoveSize"
            
        '"Quick fix" tool(s)
        Case QUICK_FIX_LIGHTING
            Load toolpanel_NDFX
            m_ActiveToolPanelKey = "NDFX"
            
        'Rectangular, Elliptical, Line selections
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
            Load toolpanel_Selections
            m_ActiveToolPanelKey = "Selections"
            
        'Vector tools
        Case VECTOR_TEXT
            Load toolpanel_Text
            m_ActiveToolPanelKey = "Text"
            
        Case VECTOR_FANCYTEXT
            Load toolpanel_FancyText
            m_ActiveToolPanelKey = "FancyText"
            
        'If a tool does not require an extra settings panel, set the active panel to -1.  This will hide all panels.
        Case Else
            m_ActiveToolPanelKey = "NoPanels"
            
    End Select
        
    'If a selection tool is active, we will also need activate a specific subpanel
    Dim activeSelectionSubpanel As Long
    If (getSelectionShapeFromCurrentTool > -1) Then
    
        activeSelectionSubpanel = Selection_Handler.getSelectionSubPanelFromCurrentTool
        
        For i = 0 To toolpanel_Selections.picSelectionSubcontainer.Count - 1
            If i = activeSelectionSubpanel Then
                toolpanel_Selections.picSelectionSubcontainer(i).Visible = True
            Else
                toolpanel_Selections.picSelectionSubcontainer(i).Visible = False
            End If
        Next i
        
    End If
    
    'Next, some tools display information about the current layer.  Synchronize that information before proceeding, so that the
    ' option panel's information is correct as soon as the window appears.
    Tool_Support.syncToolOptionsUIToCurrentLayer
    
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
            Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
            
        Else
            Process "Remove selection", , , UNDO_SELECTION
        End If
        
    End If
    
    'If the current tool is a selection tool, make sure the selection area box (interior/exterior/border) is enabled properly
    If (getSelectionShapeFromCurrentTool > -1) Then toolpanel_Selections.updateSelectionPanelLayout
    
    'Next, we can automatically hide the options toolbox for certain tools (because they have no options).  This is a
    ' nice courtesy, as it frees up space on the main canvas area if the current tool has no adjustable options.
    ' (Note that we can skip the check if the main form is not yet loaded.)
    If FormMain.Visible Then
    
        Select Case g_CurrentTool
            
            'Hand tool is currently the only tool without additional options
            Case NAV_DRAG
                g_WindowManager.SetWindowVisibility toolbar_Options.hWnd, False, False
                
            'All other tools expose options, so display the toolbox (unless the user has disabled the window completely)
            Case Else
                g_WindowManager.SetWindowVisibility toolbar_Options.hWnd, g_UserPreferences.GetPref_Boolean("Core", "Show Selections Toolbox", True), False
                
        End Select
        
    End If
    
    'Display the current tool options panel, while hiding all inactive ones
    Dim toolPanelCollection As pdDictionary
    Set toolPanelCollection = New pdDictionary
    
    If Not (toolpanel_MoveSize Is Nothing) Then toolPanelCollection.AddEntry "MoveSize", toolpanel_MoveSize.hWnd
    If Not (toolpanel_NDFX Is Nothing) Then toolPanelCollection.AddEntry "NDFX", toolpanel_NDFX.hWnd
    If Not (toolpanel_Selections Is Nothing) Then toolPanelCollection.AddEntry "Selections", toolpanel_Selections.hWnd
    If Not (toolpanel_Text Is Nothing) Then toolPanelCollection.AddEntry "Text", toolpanel_Text.hWnd
    If Not (toolpanel_FancyText Is Nothing) Then toolPanelCollection.AddEntry "FancyText", toolpanel_FancyText.hWnd
    
    g_WindowManager.DeactivateToolPanel False
    
    'To prevent flicker, we handle this in two passes.
    
    'First, activate the new window.
    If toolPanelCollection.getNumOfEntries > 0 Then
    
        For i = 0 To toolPanelCollection.getNumOfEntries - 1
            
            'If this is the active panel, display it
            If StrComp(toolPanelCollection.getKeyByIndex(i), LCase(m_ActiveToolPanelKey)) = 0 Then
                g_WindowManager.ActivateToolPanel toolPanelCollection.getValueByIndex(i)
            End If
            
        Next i
        
        'Next, hide all other panels
        For i = 0 To toolPanelCollection.getNumOfEntries - 1
            If StrComp(toolPanelCollection.getKeyByIndex(i), LCase(m_ActiveToolPanelKey)) <> 0 Then
                g_WindowManager.SetVisibilityOfAnyWindowByHwnd toolPanelCollection.getValueByIndex(i), False
            End If
        Next i
        
    End If
        
    'Display the current tool options panel, while hiding all inactive ones.  The On Error Resume statement is used to fix
    ' trouble with the .SetFocus line, below.  That .SetFocus line is helpful for fixing some VB issues with controls embedded
    ' on a picture box (specifically, combo boxes which do not drop-down properly unless a picture box or its child already
    ' has focus).  Sometimes, VB will inexplicably fail to set focus, and it will raise an Error 5 to match; as this is not
    ' a crucial error, just a VB quirk, I don't mind using OERN here.
    '
    'DISABLED PENDING ADDITIONAL TESTING WITH THE NEW PER-WINDOW OPTIONS PANEL SYSTEM (APRIL 2015)
    '
    'On Error Resume Next
    'For i = 0 To toolbar_Options.picTools.Count - 1
    '    If i = activeToolPanel Then
    '        If Not toolbar_Options.picTools(i).Visible Then
    '            toolbar_Options.picTools(i).Visible = True
    '            toolbar_Options.picTools(i).Refresh
    '            setArrowCursor toolbar_Options.picTools(i)
    '        End If
    '        If toolbar_Options.Visible And toolbar_Options.picTools(i).Visible Then toolbar_Options.picTools(i).SetFocus
    '    Else
    '        If toolbar_Options.picTools(i).Visible Then toolbar_Options.picTools(i).Visible = False
    '    End If
    'Next i
            
    newToolSelected
        
End Sub

Private Sub cmdTools_Click(Index As Integer)
        
    'If the user is dragging the mouse in from the right, and the toolbox has been shrunk from its default setting, the class cursor
    ' for forms may get stuck on the west/east "resize" cursor.  To avoid this, reset it after any button click.
    cMouseEvents.setSystemCursor IDC_ARROW
    
    'Update the previous and current tool entries
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = Index
    
    'Update the tool options area to match the newly selected tool
    resetToolButtonStates
    
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
    ReflowToolboxLayout
    
End Sub

'Used to change the display size of toolbox buttons.  newSize is expected on the range [0, 2] for small, medium, large
Public Sub updateButtonSize(ByVal newSize As Long, Optional ByVal suppressRedraw As Boolean = False)
    
    'Export the updated size to file
    If Not suppressRedraw Then g_UserPreferences.SetPref_Long "Core", "Toolbox Button Size", newSize
    
    'Update our internal size metrics to match
    m_ButtonSize = newSize
    
    Select Case m_ButtonSize
    
        Case PDTBS_SMALL
            m_ButtonWidth = FixDPI(BTS_WIDTH_SMALL)
            m_ButtonHeight = FixDPI(BTS_HEIGHT_SMALL)
        
        Case PDTBS_MEDIUM
            m_ButtonWidth = FixDPI(BTS_WIDTH_MEDIUM)
            m_ButtonHeight = FixDPI(BTS_HEIGHT_MEDIUM)
        
        Case PDTBS_LARGE
            m_ButtonWidth = FixDPI(BTS_WIDTH_LARGE)
            m_ButtonHeight = FixDPI(BTS_HEIGHT_LARGE)
    
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
    
    g_WindowManager.UpdateMinimumDimensions Me.hWnd, m_ButtonWidth
    
    'Reflow the interface as requested
    If Not suppressRedraw Then ReflowToolboxLayout
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) MakeFormPretty is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()

    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    MakeFormPretty Me
    
    'Tooltips must be manually re-assigned according to the current language.  This is a necessary evil, if the user switches
    ' between two non-English languages at run-time.
    
    'File tool buttons come first
    cmdFile(FILE_NEW).AssignTooltip "This option will create a blank image.  Other ways to create new images can be found in the File -> Import menu.", "New Image"
    cmdFile(FILE_OPEN).AssignTooltip "Another way to open images is dragging them from your desktop or Windows Explorer and dropping them onto PhotoDemon.", "Open one or more images for editing"
    
    If g_ConfirmClosingUnsaved Then
        cmdFile(FILE_CLOSE).AssignTooltip "If the current image has not been saved, you will receive a prompt to save it before it closes.", "Close the current image"
    Else
        cmdFile(FILE_CLOSE).AssignTooltip "Because you have turned off save prompts (via Edit -> Preferences), you WILL NOT receive a prompt to save this image before it closes.", "Close the current image"
    End If
    
    If g_UserPreferences.GetPref_Long("Saving", "Overwrite Or Copy", 0) = 0 Then
        cmdFile(FILE_SAVE).AssignTooltip "WARNING: this will overwrite the current image file.  To save to a different file, use the ""Save As"" button.", "Save image in current format"
    Else
        cmdFile(FILE_SAVE).AssignTooltip "You have specified ""safe"" save mode, which means that each save will create a new file with an auto-incremented filename.", "Save image in current format"
    End If
    
    cmdFile(FILE_SAVEAS_LAYERS).AssignTooltip "Use this to quickly save a lossless copy of the current image.  The lossless copy will be saved in PDI format, in the image's current folder, using the current filename (plus an auto-incremented number, as necessary).", "Save lossless copy"
    cmdFile(FILE_SAVEAS_FLAT).AssignTooltip "The Save As command always raises a dialog, so you can specify a new file name, folder, and/or image format for the current image.", "Save As (export to new format or filename)"
        
    'Non-destructive tool buttons are next
    cmdTools(NAV_DRAG).AssignTooltip "Hand (click-and-drag image scrolling)"
    cmdTools(NAV_MOVE).AssignTooltip "Move and resize image layers"
    cmdTools(QUICK_FIX_LIGHTING).AssignTooltip "Apply non-destructive lighting adjustments"
    
    '...then selections...
    cmdTools(SELECT_RECT).AssignTooltip "Rectangular Selection"
    cmdTools(SELECT_CIRC).AssignTooltip "Elliptical (Oval) Selection"
    cmdTools(SELECT_LINE).AssignTooltip "Line Selection"
    cmdTools(SELECT_POLYGON).AssignTooltip "Polygon Selection"
    cmdTools(SELECT_LASSO).AssignTooltip "Lasso (Freehand) Selection"
    cmdTools(SELECT_WAND).AssignTooltip "Magic Wand Selection"
    
    '...then vector tools...
    cmdTools(VECTOR_TEXT).AssignTooltip "Text (basic)"
    cmdTools(VECTOR_FANCYTEXT).AssignTooltip "Typography (advanced)"
    
    'The right separator line is colored according to the current shadow accent color
    If Not g_Themer Is Nothing Then
        lnRightSeparator.borderColor = g_Themer.GetThemeColor(PDTC_GRAY_SHADOW)
    Else
        lnRightSeparator.borderColor = vbHighlight
    End If
    
End Sub
