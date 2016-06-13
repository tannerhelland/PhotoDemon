VERSION 5.00
Begin VB.Form toolbar_Toolbox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "File"
   ClientHeight    =   9810
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   2115
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   654
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   141
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   30
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
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
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   1620
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "undo"
   End
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   2580
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "n-d"
   End
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   25
      Top             =   3540
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "select"
   End
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   26
      Top             =   5100
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "text"
   End
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   27
      Top             =   6120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "paint"
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
'Copyright 2013-2016 by Tanner Helland
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

'When toggling tool-buttons, we set a module-level check to ensure that each toggle doesn't cause us to
' re-enter the reflow code.
Private m_InsideReflowCode As Boolean

'This form supports a variety of resize modes, and we use cMouseEvents to handle cursor duties
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

Private Sub cmdFile_Click(Index As Integer)
        
    'If the user is dragging the mouse in from the right, and the toolbox has been shrunk from its default setting, the class cursor
    ' for forms may get stuck on the west/east "resize" cursor.  To avoid this, reset it after any button click.
    cMouseEvents.SetSystemCursor IDC_ARROW
        
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
    cMouseEvents.SetSystemCursor IDC_ARROW
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
        cMouseEvents.SetSystemCursor IDC_SIZEWE
        
        If (Button = vbLeftButton) Then
            m_WeAreResponsibleForResize = True
            ReleaseCapture
            SendMessage Me.hWnd, WM_NCLBUTTONDOWN, hitCode, ByVal 0&
            
            'After the toolbox has been resized, we need to manually notify the toolbox manager, so it can
            ' notify any neighboring toolboxes (and/or the central canvas)
            Toolboxes.SetConstrainingSize PDT_LeftToolbox, Me.ScaleWidth
            FormMain.UpdateMainLayout
            
            'A premature exit is required, because the end of this sub contains code to detect the release of the
            ' mouse after a drag event.  Because the event is not being initiated normally, we can't detect a standard
            ' MouseUp event, so instead, we mimic it by checking MouseMove and m_WeAreResponsibleForResize = TRUE.
            Exit Sub
            
        End If
        
    Else
        cMouseEvents.SetSystemCursor IDC_ARROW
    End If
    
    'Check for mouse release; we will only reach this point if the mouse is *not* in resize territory, which in turn
    ' means we can free the release code and resize the window now.  (On some OS/theme combinations, the canvas will
    ' live-resize as the mouse is moved.  On others, the canvas won't redraw until the mouse is released.)
    If m_WeAreResponsibleForResize Then
        
        m_WeAreResponsibleForResize = False
        
        'TODO: make sure this is okay with 7.0's new toolbox manager
        
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
    cMouseEvents.AddInputTracker Me.hWnd, True, True, , True
        
    g_PreviousTool = -1
    g_CurrentTool = g_UserPreferences.GetPref_Long("Tools", "LastUsedTool", NAV_DRAG)
    
    'Retrieve any relevant toolbox display settings from the user's preferences file
    m_ShowCategoryLabels = g_UserPreferences.GetPref_Boolean("Core", "Show Toolbox Category Labels", True)
    m_ButtonSize = g_UserPreferences.GetPref_Long("Core", "Toolbox Button Size", 1)
    
    'Note that we don't actually reflow the interface here; that will happen later, when the form's previous size and
    ' position are loaded from the user's preference file.
    
    'As a final step, redraw everything against the current theme.
    UpdateAgainstCurrentTheme
    
End Sub

Private Sub Form_LostFocus()
    cMouseEvents.SetSystemCursor IDC_DEFAULT
End Sub

'Reflow the form's contents
Private Sub Form_Resize()
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
    For i = 0 To ttlCategories.UBound
        ttlCategories(i).SetWidth m_rightBoundary - (ttlCategories(i).GetLeft)
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
    For i = 0 To ttlCategories.UBound
        ttlCategories(i).Visible = m_ShowCategoryLabels
    Next i
    
    'File group.  We position this label manually, and it serves as the reference for all subsequent labels.
    If m_ShowCategoryLabels Then
        ttlCategories(0).SetLeft m_hOffsetDefaultLabel
        ttlCategories(0).SetTop FixDPI(2)
        vOffset = ttlCategories(0).GetTop + ttlCategories(0).GetHeight + m_labelMarginBottom
    Else
        vOffset = FixDPI(2)
    End If
    
    ReflowButtonSet 0, False, FILE_NEW, FILE_SAVEAS_FLAT, hOffset, vOffset
    
    'Undo group
    PositionToolLabel 1, cmdFile(FILE_SAVEAS_FLAT), hOffset, vOffset
    ReflowButtonSet 1, False, FILE_UNDO, FILE_REDO, hOffset, vOffset
    
    'Non-destructive group
    PositionToolLabel 2, cmdFile(FILE_REDO), hOffset, vOffset
    ReflowButtonSet 2, True, NAV_DRAG, QUICK_FIX_LIGHTING, hOffset, vOffset
    
    'Selection group
    PositionToolLabel 3, cmdTools(QUICK_FIX_LIGHTING), hOffset, vOffset
    ReflowButtonSet 3, True, SELECT_RECT, SELECT_WAND, hOffset, vOffset
    
    'Vector group
    PositionToolLabel 4, cmdTools(SELECT_WAND), hOffset, vOffset
    ReflowButtonSet 4, True, VECTOR_TEXT, VECTOR_FANCYTEXT, hOffset, vOffset
    
    'Paint group
    PositionToolLabel 5, cmdTools(VECTOR_FANCYTEXT), hOffset, vOffset
    ReflowButtonSet 5, True, PAINT_BASICBRUSH, PAINT_BASICBRUSH, hOffset, vOffset
        
    'Macro recording message
    If (vOffset < cmdTools(cmdTools.UBound).Top + cmdTools(cmdTools.UBound).Height) Then
        vOffset = cmdTools(cmdTools.UBound).Top + cmdTools(cmdTools.UBound).Height + m_buttonMarginBottom
    End If
    
    vOffset = vOffset + m_labelMarginTop
    lblRecording.Move lblRecording.Left, vOffset

End Sub

'Companion function to ReflowToolboxLayout(), above.
Private Sub PositionToolLabel(ByRef targetLabelIndex As Long, ByRef referenceButton As Object, ByRef hOffset As Long, ByRef vOffset As Long)
    
    If m_ShowCategoryLabels Then
        
        Dim heightCalc As Long
        If ttlCategories(targetLabelIndex - 1).Value Then heightCalc = referenceButton.Height Else heightCalc = 0
        
        'If vOffset < referenceButton.Top + heightCalc Then
            vOffset = referenceButton.Top + heightCalc + m_buttonMarginBottom
        'End If
        
        vOffset = vOffset + m_labelMarginTop
        ttlCategories(targetLabelIndex).SetLeft m_hOffsetDefaultLabel
        ttlCategories(targetLabelIndex).SetTop vOffset
        vOffset = ttlCategories(targetLabelIndex).GetTop + ttlCategories(targetLabelIndex).GetHeight + m_labelMarginBottom
        hOffset = m_hOffsetDefaultButton
        
    End If
    
End Sub

'Companion function to ReflowToolboxLayout(), above.
Private Sub ReflowButtonSet(ByVal associatedTitleIndex As Long, ByVal categoryIsTools As Boolean, ByVal startIndex As Long, ByVal endIndex As Long, ByRef hOffset As Long, ByRef vOffset As Long)
    
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
        
        'If the associated title index is set to TRUE, display the button and calculate a new offset for the next button
        targetObject.Visible = ttlCategories(associatedTitleIndex).Value
        
        If ttlCategories(associatedTitleIndex).Value Then
            hOffset = hOffset + targetObject.Width + m_buttonMarginRight
            If hOffset + targetObject.Width > m_rightBoundary Then
                hOffset = m_hOffsetDefaultButton
                vOffset = vOffset + targetObject.Height + m_buttonMarginBottom
            End If
        End If
        
    Next i
    
End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the ToggleToolboxVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    g_UserPreferences.SetPref_Long "Tools", "LastUsedTool", g_CurrentTool
    If g_ProgramShuttingDown Then ReleaseFormTheming Me
End Sub

'When a new tool is selected, we may need to initialize certain values.
Private Sub NewToolSelected()
    
    Select Case g_CurrentTool
    
        'Selection tools
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO
        
            'See if a selection is already active on the image
            If SelectionsAllowed(False) Then
            
                'A selection is already active!
                
                'If the existing selection type matches the tool type, no problem - activate the transform tools
                ' (if relevant), but make no other changes to the image
                If (g_CurrentTool = Selection_Handler.getRelevantToolFromSelectShape()) Then
                    SetUIGroupState PDUI_SelectionTransforms, pdImages(g_CurrentImage).mainSelection.isTransformable
                
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
            If (g_OpenImageCount > 0) Then Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                
        Case VECTOR_TEXT, VECTOR_FANCYTEXT
        
            'Switching text tools may require us to redraw the text buffer, as the font rendering engine changes depending
            ' on the current text tool.
            If (g_OpenImageCount > 0) Then Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        Case Else
        
            'Finally, because tools may do some custom rendering atop the image canvas, now is a good time to redraw the canvas.
            ' (Note that we can use a very late pipeline stage, as only tool-specific overlays need to be redrawn.)
            If (g_OpenImageCount > 0) Then Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
    End Select
        
End Sub

'External functions can use this to request the selection of a new tool (for example, Select All uses this to set the
' rectangular tool selector as the current tool)
Public Sub SelectNewTool(ByVal newToolID As PDTools)
    
    If newToolID <> g_CurrentTool Then
        g_PreviousTool = g_CurrentTool
        g_CurrentTool = newToolID
        ResetToolButtonStates
    End If
    
End Sub

'When a new tool button is selected, we need to raise all the others and display the proper options box
Public Sub ResetToolButtonStates()
        
    m_InsideReflowCode = True
        
    'Start by depressing the selected button and raising all unselected ones
    Dim catID As Long
    For catID = 0 To cmdTools.Count - 1
        If catID = g_CurrentTool Then
            If Not cmdTools(catID).Value Then cmdTools(catID).Value = True
        Else
            If cmdTools(catID).Value Then cmdTools(catID).Value = False
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
            toolpanel_MoveSize.UpdateAgainstCurrentTheme
            m_ActiveToolPanelKey = "MoveSize"
            
        '"Quick fix" tool(s)
        Case QUICK_FIX_LIGHTING
            Load toolpanel_NDFX
            toolpanel_NDFX.UpdateAgainstCurrentTheme
            m_ActiveToolPanelKey = "NDFX"
            
        'Rectangular, Elliptical, Line selections
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
            Load toolpanel_Selections
            toolpanel_Selections.UpdateAgainstCurrentTheme
            m_ActiveToolPanelKey = "Selections"
            
        'Vector tools
        Case VECTOR_TEXT
            Load toolpanel_Text
            toolpanel_Text.UpdateAgainstCurrentTheme
            m_ActiveToolPanelKey = "Text"
            
        Case VECTOR_FANCYTEXT
            Load toolpanel_FancyText
            toolpanel_FancyText.UpdateAgainstCurrentTheme
            m_ActiveToolPanelKey = "FancyText"
            
        'If a tool does not require an extra settings panel, set the active panel to -1.  This will hide all panels.
        Case Else
            m_ActiveToolPanelKey = "NoPanels"
            
    End Select
        
    'If a selection tool is active, we will also need activate a specific subpanel
    Dim activeSelectionSubpanel As Long
    If (getSelectionShapeFromCurrentTool > -1) Then
    
        activeSelectionSubpanel = Selection_Handler.GetSelectionSubPanelFromCurrentTool
        
        For i = 0 To toolpanel_Selections.ctlGroupSelectionSubcontainer.Count - 1
            toolpanel_Selections.ctlGroupSelectionSubcontainer(i).Visible = CBool(i = activeSelectionSubpanel)
        Next i
        
    End If
    
    'Next, some tools display information about the current layer.  Synchronize that information before proceeding, so that the
    ' option panel's information is correct as soon as the window appears.
    Tool_Support.SyncToolOptionsUIToCurrentLayer
    
    'Check the selection state before swapping tools.  If a selection is active, and the user is switching to the same
    ' tool used to create the current selection, we don't want to erase the current selection.  If they are switching
    ' to a *different* selection tool, however, then we *do* want to erase the current selection.
    If SelectionsAllowed(False) And (getRelevantToolFromSelectShape() <> g_CurrentTool) And (getSelectionShapeFromCurrentTool > -1) Then
        
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
    If (getSelectionShapeFromCurrentTool > -1) Then toolpanel_Selections.UpdateSelectionPanelLayout
    
    'Next, we can automatically hide the options toolbox for certain tools (because they have no options).  This is a
    ' nice courtesy, as it frees up space on the main canvas area if the current tool has no adjustable options.
    ' (Note that we can skip the check if the main form is not yet loaded.)
    If FormMain.Visible Then
    
        Select Case g_CurrentTool
            
            'Hand tool is currently the only tool without additional options
            Case NAV_DRAG
                Toolboxes.SetToolboxVisibility PDT_BottomToolbox, False
                
            'All other tools expose options, so display the toolbox (unless the user has disabled the window completely)
            Case Else
                Toolboxes.SetToolboxVisibilityByPreference PDT_BottomToolbox
                
        End Select
        
        FormMain.UpdateMainLayout
        
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
    If (toolPanelCollection.GetNumOfEntries > 0) Then
    
        For i = 0 To toolPanelCollection.GetNumOfEntries - 1
            
            'If this is the active panel, display it
            If (StrComp(toolPanelCollection.GetKeyByIndex(i), LCase(m_ActiveToolPanelKey)) = 0) Then
                g_WindowManager.ActivateToolPanel toolPanelCollection.GetValueByIndex(i), toolbar_Options.hWnd
            End If
            
        Next i
        
        'Next, hide all other panels
        For i = 0 To toolPanelCollection.GetNumOfEntries - 1
            If (StrComp(toolPanelCollection.GetKeyByIndex(i), LCase(m_ActiveToolPanelKey)) <> 0) Then
                g_WindowManager.SetVisibilityByHWnd toolPanelCollection.GetValueByIndex(i), False
            End If
        Next i
        
    End If
            
    NewToolSelected
    
    m_InsideReflowCode = False
        
End Sub

Private Sub cmdTools_Click(Index As Integer)
    
    'Update the previous and current tool entries
    If cmdTools(Index).Value Then
        g_PreviousTool = g_CurrentTool
        g_CurrentTool = Index
    End If
    
    'Update the tool options area to match the newly selected tool
    If (Not m_InsideReflowCode) And (cmdTools(Index).Value) Then
        
        'If the user is dragging the mouse in from the right, and the toolbox has been shrunk from its default setting, the class cursor
        ' for forms may get stuck on the west/east "resize" cursor.  To avoid this, reset it after any button click.
        cMouseEvents.SetSystemCursor IDC_ARROW
        
        'Repaint all tool buttons to reflect the new selection
        ResetToolButtonStates
    
    End If
    
End Sub

'Used to change the visibility of category labels.  When disabled, the button layout is reflowed into a continuous
' stream of buttons.  When enabled, buttons are sorted by category.
Public Sub ToggleToolCategoryLabels()
    
    m_ShowCategoryLabels = CBool(Not m_ShowCategoryLabels)
    FormMain.MnuWindowToolbox(2).Checked = m_ShowCategoryLabels
    g_UserPreferences.SetPref_Boolean "Core", "Show Toolbox Category Labels", m_ShowCategoryLabels
    
    'Reflow the interface
    ReflowToolboxLayout
    
End Sub

'Used to change the display size of toolbox buttons.  newSize is expected on the range [0, 2] for small, medium, large
Public Sub UpdateButtonSize(ByVal newSize As Long, Optional ByVal suppressRedraw As Boolean = False)
    
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
    
    'TODO: notify the new toolbox manager of this change
    'g_WindowManager.UpdateMinimumDimensions Me.hWnd, m_ButtonWidth
    
    'Reflow the interface as requested
    If Not suppressRedraw Then ReflowToolboxLayout
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()

    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me
    
    'Tooltips must be manually re-assigned according to the current language.  This is a necessary evil, if the user switches
    ' between two non-English languages at run-time.
    
    'Title bars first
    ttlCategories(0).AssignTooltip "File tools: create, load, and save image files"
    ttlCategories(1).AssignTooltip "Undo/Redo tools: revert destructive changes made to an image or layer"
    ttlCategories(2).AssignTooltip "Non-destructive tools: move, rotate, or apply certain photo adjustments without permanently modifying an image"
    ttlCategories(3).AssignTooltip "Select tools: isolate parts of an image or layer for further editing"
    ttlCategories(4).AssignTooltip "Text tools: create and edit text layers"
    ttlCategories(5).AssignTooltip "Paint tools: use a mouse, touchpad, or pen tablet to apply brushstrokes to an image or layer"
    
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
        lnRightSeparator.borderColor = g_Themer.GetGenericUIColor(UI_GrayDark)
    Else
        lnRightSeparator.borderColor = vbHighlight
    End If
    
End Sub

Private Sub ttlCategories_Click(Index As Integer, ByVal newState As Boolean)
    ReflowToolboxLayout
End Sub
