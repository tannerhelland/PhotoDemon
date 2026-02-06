VERSION 5.00
Begin VB.Form toolbar_Toolbox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "File"
   ClientHeight    =   9810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
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
   Icon            =   "Toolbar_ToolSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   654
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   30
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "file"
      CustomDragDropEnabled=   -1  'True
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
      CustomDragDropEnabled=   -1  'True
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
      CustomDragDropEnabled=   -1  'True
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
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   6
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   7
      Left            =   840
      TabIndex        =   5
      Top             =   4440
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   8
      Left            =   1560
      TabIndex        =   6
      Top             =   4440
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   9
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   10
      Left            =   840
      TabIndex        =   8
      Top             =   5040
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   300
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
      AutoToggle      =   -1  'True
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   1
      Left            =   840
      TabIndex        =   10
      Top             =   300
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
      AutoToggle      =   -1  'True
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   2
      Left            =   1560
      TabIndex        =   11
      Top             =   300
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
      AutoToggle      =   -1  'True
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
      AutoToggle      =   -1  'True
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   4
      Left            =   840
      TabIndex        =   13
      Top             =   960
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
      AutoToggle      =   -1  'True
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   5
      Left            =   1560
      TabIndex        =   14
      Top             =   960
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
      AutoToggle      =   -1  'True
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
      AutoToggle      =   -1  'True
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdFile 
      Height          =   600
      Index           =   7
      Left            =   840
      TabIndex        =   16
      Top             =   1920
      Width           =   720
      _ExtentX        =   1085
      _ExtentY        =   1085
      AutoToggle      =   -1  'True
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   11
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   12
      Left            =   840
      TabIndex        =   18
      Top             =   6000
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   13
      Left            =   120
      TabIndex        =   19
      Top             =   7080
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   1620
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "undo"
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   2580
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "layout"
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   4140
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "select"
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   5700
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "text"
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdTitle ttlCategories 
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   6720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "paint"
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   14
      Left            =   840
      TabIndex        =   0
      Top             =   7080
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblRecording 
      Height          =   720
      Left            =   120
      Top             =   8640
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1270
      Alignment       =   2
      Caption         =   "macro recording in progress..."
      CustomDragDropEnabled=   -1  'True
      Layout          =   1
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   15
      Left            =   1560
      TabIndex        =   26
      Top             =   7080
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   16
      Left            =   120
      TabIndex        =   27
      Top             =   7680
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   3480
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   17
      Left            =   840
      TabIndex        =   29
      Top             =   7680
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   18
      Left            =   1560
      TabIndex        =   30
      Top             =   7680
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   4
      Left            =   840
      TabIndex        =   31
      Top             =   3480
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdTools 
      Height          =   600
      Index           =   5
      Left            =   1560
      TabIndex        =   32
      Top             =   3480
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      CustomDragDropEnabled=   -1  'True
   End
End
Attribute VB_Name = "toolbar_Toolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Primary Toolbox
'Copyright 2013-2026 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 06/March/22
'Last update: rename "non-destructive" tool group to "layout tools"
'
'This form was initially integrated into the main MDI form.  In fall 2013, PhotoDemon left behind
' the MDI model, and all toolboxes were moved to their own windows.  This toolbox now manages
' all on-canvas tools, while also providing shortcuts to open/save/undo/redo tasks.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Because this form is resizable at run-time, we need to play some games with mouse capturing
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTRIGHT As Long = 11

'The value of all controls on this form are saved and loaded to file by this class
' (Normally this is declared WithEvents, but this dialog doesn't require custom settings behavior.)
Private m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'When we are responsible for this window resizing (because the user is resizing our window manually), we set this to TRUE.
' This variable is then checked before requesting additional redraws during our resize event.
Private m_WeAreResponsibleForResize As Boolean

'The user has various options for how this toolbox is displayed.  When changed, those options need to update these
' module-level values, then manually reflow the interface.
Private m_ShowCategoryLabels As Boolean

'Button size is also controllable.  In the future, this will result in an actual change to the images used on the buttons.
' For now, however, we simply resize the buttons themselves.
Public Enum PD_ToolboxButtonSize
    tbs_Small = 0
    tbs_Medium = 1
    tbs_Large = 2
End Enum

#If False Then
    Private Const tbs_Small = 0, tbs_Medium = 1, tbs_Large = 2
#End If

Private m_ButtonSize As PD_ToolboxButtonSize

Private Const BTS_WIDTH_SMALL As Long = 32
Private Const BTS_WIDTH_MEDIUM As Long = 46
Private Const BTS_WIDTH_LARGE As Long = 56

Private Const BTS_HEIGHT_SMALL As Long = 32
Private Const BTS_HEIGHT_MEDIUM As Long = 38
Private Const BTS_HEIGHT_LARGE As Long = 48

Private Const BTS_IMG_SMALL As Long = 18
Private Const BTS_IMG_MEDIUM As Long = 24
Private Const BTS_IMG_LARGE As Long = 36

Private m_ButtonWidth As Long, m_ButtonHeight As Long

'These values are basically constants; they are set by the ReflowToolboxLayout function.
Private m_hOffsetDefaultLabel As Long, m_hOffsetDefaultButton As Long
Private m_labelMarginBottom As Long, m_labelMarginTop As Long
Private m_buttonMarginBottom As Long, m_buttonMarginRight As Long
Private m_rightBoundary As Long

'When toggling tool-buttons, we set a module-level check to ensure that each toggle doesn't cause us to
' re-enter the reflow code.
Private m_InsideReflowCode As Boolean

'This form supports a variety of resize modes, and we use m_MouseEvents to handle cursor duties
Private WithEvents m_MouseEvents As pdInputMouse
Attribute m_MouseEvents.VB_VarHelpID = -1

'This form also manages individual toolpanel windows (because they are shown/hidden based upon interactions
' with the buttons on *this* form).
Private Type ToolPanelTracker
    PanelHWnd As Long
    PanelWasLoaded As Boolean
End Type

Private m_NumOfPanels As Long
Private m_Panels() As ToolPanelTracker

'Each individual tool panel must have a unique entry inside *this* enum.  Note that a number of
' tools share panels, so this number has no meaningful relation to the net number of tools available.
Private Enum PD_ToolPanels
    TP_None = -1
    TP_MoveSize = 0
    TP_ColorPicker = 1
    TP_Measure = 2
    TP_Crop = 3
    
    TP_Selections = 4
    
    TP_Text = 5
    TP_Typography = 6
    
    TP_Pencil = 7
    TP_Paintbrush = 8
    TP_Eraser = 9
    TP_Clone = 10
    TP_Fill = 11
    TP_Gradient = 12
End Enum

Private Const NUM_OF_TOOL_PANELS As Long = TP_Gradient + 1

#If False Then
    Private Const TP_None = -1, TP_MoveSize = 0, TP_ColorPicker = 1, TP_Measure = 2, TP_Crop = 3
    Private Const TP_Selections = 4
    Private Const TP_Text = 5, TP_Typography = 6
    Private Const TP_Pencil = 7, TP_Paintbrush = 8, TP_Eraser = 9, TP_Clone = 10, TP_Fill = 11, TP_Gradient = 12
#End If

'The currently active tool panel will be mirrored to this value
Private m_ActiveToolPanel As PD_ToolPanels

'Sometimes, external functions need to get a list of valid tool names.  (The search bar, for example.)
' We store a list of localized tool names and corresponding action strings internally.
Private Type PD_ToolboxAction
    ta_ToolName As String
    ta_ToolAction As String
End Type

Private m_ToolActions() As PD_ToolboxAction, m_numToolActions As Long

'hWnd of the currently active tool panel.  We must track this and pass it to the window manager
' when loading a new tool panel.
Private m_CurrentToolPanelHwnd As Long

Private Sub cmdFile_Click(Index As Integer, ByVal Shift As ShiftConstants)
        
    'If the user is dragging the mouse in from the right, and the toolbox has been shrunk from its default setting, the class cursor
    ' for forms may get stuck on the west/east "resize" cursor.  To avoid this, reset it after any button click.
    m_MouseEvents.SetCursor_System IDC_ARROW
        
    Select Case Index
    
        Case FILE_NEW
            Actions.LaunchAction_ByName "file_new"
            
        Case FILE_OPEN
            Actions.LaunchAction_ByName "file_open"
            
        Case FILE_CLOSE
            Actions.LaunchAction_ByName "file_close"
        
        Case FILE_SAVE
            Actions.LaunchAction_ByName "file_save"
        
        Case FILE_SAVEAS_LAYERS
            Actions.LaunchAction_ByName "file_savecopy"
            
        Case FILE_SAVEAS_FLAT
            Actions.LaunchAction_ByName "file_saveas"
        
        Case FILE_UNDO
            Actions.LaunchAction_ByName "edit_undo"
            
        Case FILE_REDO
            Actions.LaunchAction_ByName "edit_redo"
    
    End Select
    
End Sub

Private Sub cmdFile_CustomDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Loading.LoadFromDragDrop Data, Effect, Button, Shift
End Sub

Private Sub cmdFile_CustomDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub cmdTools_CustomDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Loading.LoadFromDragDrop Data, Effect, Button, Shift
End Sub

Private Sub cmdTools_CustomDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Loading.LoadFromDragDrop Data, Effect, Button, Shift
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    If Not (m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
    End If
    
End Sub

Private Sub lblRecording_CustomDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Loading.LoadFromDragDrop Data, Effect, Button, Shift
End Sub

Private Sub lblRecording_CustomDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub

'When the mouse leaves this toolbox, reset it to an arrow (so other forms don't magically acquire the west/east resize cursor, as the mouse is
' likely to leave off the right side of this form)
Private Sub m_MouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseEvents.SetCursor_System IDC_ARROW
End Sub

Private Sub m_MouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'Ignore all mouse events while the user is interacting with the canvas
    If FormMain.MainCanvas(0).IsMouseDown(pdLeftButton Or pdRightButton) Then Exit Sub
    
    'If the mouse is near the resizable edge of the toolbar (the left edge, currently), allow the user to resize
    ' the layer toolbox.
    Dim mouseInResizeTerritory As Boolean
    
    'How close does the mouse have to be to the form border to allow resizing; currently we use 7 pixels, while accounting
    ' for DPI variance (e.g. 7 pixels at 96 dpi)
    Dim resizeBorderAllowance As Long
    resizeBorderAllowance = Interface.FixDPI(7)
    
    Dim hitCode As Long
    
    'Check the mouse position to see if it's in resize territory (along the left edge of the toolbox)
    If (y > 0) And (y < Me.ScaleHeight) And (x > (Me.ScaleWidth - resizeBorderAllowance)) Then
        mouseInResizeTerritory = True
        hitCode = HTRIGHT
    End If
    
    'If the left mouse button is down, and the mouse is in resize territory, initiate an API resize event
    If mouseInResizeTerritory Then
    
        'Change the cursor to a resize cursor
        m_MouseEvents.SetCursor_System IDC_SIZEWE
        
        If ((Button And vbLeftButton) <> 0) Then
        
            m_WeAreResponsibleForResize = True
            ReleaseCapture
            VBHacks.SendMsgW Me.hWnd, WM_NCLBUTTONDOWN, hitCode, 0&
            
            'After the toolbox has been resized, we need to manually notify the toolbox manager, so it can
            ' notify any neighboring toolboxes (and/or the central canvas)
            Toolboxes.SetConstrainingSize PDT_LeftToolbox, g_WindowManager.GetClientWidth(Me.hWnd)
            FormMain.UpdateMainLayout
            
            'A premature exit is required, because the end of this sub contains code to detect the release of the
            ' mouse after a drag event.  Because the event is not being initiated normally, we can't detect a standard
            ' MouseUp event, so instead, we mimic it by checking MouseMove and m_WeAreResponsibleForResize = TRUE.
            Exit Sub
            
        End If
        
    Else
        m_MouseEvents.SetCursor_System IDC_ARROW
    End If
    
    'Check for mouse release; we will only reach this point if the mouse is *not* in resize territory, which in turn
    ' means we can free the release code and resize the window now.  (On some OS/theme combinations, the canvas will
    ' live-resize as the mouse is moved.  On others, the canvas won't redraw until the mouse is released.)
    If m_WeAreResponsibleForResize Then m_WeAreResponsibleForResize = False
    
End Sub

Private Sub Form_Load()
    
    'Retrieve any relevant toolbox display settings from the user's preferences file
    m_ShowCategoryLabels = UserPrefs.GetPref_Boolean("Core", "Show Toolbox Category Labels", True)
    m_ButtonSize = UserPrefs.GetPref_Long("Core", "Toolbox Button Size", tbs_Small)
    
    'Reset a stack of tool names, actions, and associated hotkeys
    m_numToolActions = 0
    ReDim m_ToolActions(0) As PD_ToolboxAction
    
    'Initialize a mouse handler
    Set m_MouseEvents = New pdInputMouse
    m_MouseEvents.AddInputTracker Me.hWnd
        
    g_PreviousTool = TOOL_UNDEFINED
    g_CurrentTool = UserPrefs.GetPref_Long("Tools", "LastUsedTool", NAV_DRAG)
    
    'Note that we don't actually reflow the interface here; that will happen later, when the form's previous size and
    ' position are loaded from the user's preference file.
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
    'As a final step, redraw everything against the current theme.
    UpdateAgainstCurrentTheme
    
End Sub

Private Sub Form_LostFocus()
    m_MouseEvents.SetCursor_System IDC_DEFAULT
End Sub

'Reflow the form's contents
Private Sub Form_Resize()
    ReflowToolboxLayout
End Sub

Private Sub ReflowToolboxLayout()
    
    'Failsafe check for some module-level properties being loaded
    If (m_ButtonWidth = 0) Then Exit Sub
    
    Dim continueWithReflow As Boolean
    
    Do
    
        continueWithReflow = True
    
        'Mark the right boundary for images, which allows for some padding when reflowing the interface
        Dim rightBoundPadding As Long
        rightBoundPadding = Interface.FixDPI(3)
        m_rightBoundary = g_WindowManager.GetClientWidth(Me.hWnd) - rightBoundPadding
        
        'Establish default positioning values (some of which are dependent on screen DPI)
        m_hOffsetDefaultLabel = Interface.FixDPI(4) 'Left-position of labels
        m_hOffsetDefaultButton = 0        'Left-position of left-most buttons
        m_labelMarginBottom = Interface.FixDPI(4)   'Distance between the bottom of labels and the top of buttons
        m_labelMarginTop = Interface.FixDPI(2)      'Distance between the bottom of buttons and top of labels
        m_buttonMarginBottom = 0          'Distance between button rows
        m_buttonMarginRight = 0           'Distance between buttons
        
        'With all default values correctly calculated, we now want to ensure that the underlying form is a nice match
        ' for our current button size.  Said another way, we want to force the toolbox's width to a clean multiple of
        ' the current button size. (This minimizes dead space on the bar's right margin.)
        Dim newWidth As Long
        newWidth = g_WindowManager.GetClientWidth(Me.hWnd) - (m_hOffsetDefaultButton + rightBoundPadding) + 1
        newWidth = Int(newWidth \ m_ButtonWidth) * m_ButtonWidth
        newWidth = newWidth + (m_hOffsetDefaultButton + rightBoundPadding)
        
        'If our calculated size differs from the actual size, resize immediately, and refresh the surrounding
        ' client area to match.  (Note that we skip a few cases - specifically, if this resize is invalid because
        ' it's too small, or if our newly calculated size matches our previously calculated size.)
        If (newWidth <> g_WindowManager.GetClientWidth(Me.hWnd)) And (newWidth > Toolboxes.GetToolboxMinWidth(PDT_LeftToolbox)) Then
            Toolboxes.SetConstrainingSize PDT_LeftToolbox, newWidth
            g_WindowManager.SetSizeByHWnd Me.hWnd, newWidth, g_WindowManager.GetClientHeight(Me.hWnd), True
            FormMain.UpdateMainLayout False
            continueWithReflow = False
        End If
        
    Loop While (Not continueWithReflow)
    
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
        
    'Reflow label width first; they are easy because they simply match the width of the form
    Dim i As Long
    For i = 0 To ttlCategories.UBound
        ttlCategories(i).SetWidth m_rightBoundary - (ttlCategories(i).GetLeft)
    Next i
    
    lblRecording.SetWidth m_rightBoundary - (lblRecording.GetLeft + FixDPI(2))
        
    'If category labels are displayed, make them visible now
    For i = 0 To ttlCategories.UBound
        ttlCategories(i).Visible = m_ShowCategoryLabels
    Next i
    
    'File group.  We position this label manually, and it serves as the reference for all subsequent labels.
    If m_ShowCategoryLabels Then
        ttlCategories(0).SetLeft m_hOffsetDefaultLabel
        ttlCategories(0).SetTop Interface.FixDPI(2)
        vOffset = ttlCategories(0).GetTop + ttlCategories(0).GetHeight + m_labelMarginBottom
    Else
        vOffset = Interface.FixDPI(2)
    End If
    
    ReflowButtonSet 0, False, FILE_NEW, FILE_SAVEAS_FLAT, hOffset, vOffset
    
    'Undo group
    PositionToolLabel 1, cmdFile(FILE_SAVEAS_FLAT), hOffset, vOffset
    ReflowButtonSet 1, False, FILE_UNDO, FILE_REDO, hOffset, vOffset
    
    'Layout group
    PositionToolLabel 2, cmdFile(FILE_REDO), hOffset, vOffset
    ReflowButtonSet 2, True, NAV_DRAG, ND_CROP, hOffset, vOffset
    
    'Selection group
    PositionToolLabel 3, cmdTools(ND_CROP), hOffset, vOffset
    ReflowButtonSet 3, True, SELECT_RECT, SELECT_WAND, hOffset, vOffset
    
    'Vector group
    PositionToolLabel 4, cmdTools(SELECT_WAND), hOffset, vOffset
    ReflowButtonSet 4, True, TEXT_BASIC, TEXT_ADVANCED, hOffset, vOffset
    
    'Paint group
    PositionToolLabel 5, cmdTools(TEXT_ADVANCED), hOffset, vOffset
    ReflowButtonSet 5, True, PAINT_PENCIL, PAINT_GRADIENT, hOffset, vOffset
    
    'Macro recording message
    If (vOffset < cmdTools(cmdTools.UBound).GetTop + cmdTools(cmdTools.UBound).GetHeight) Then
        vOffset = cmdTools(cmdTools.UBound).GetTop + cmdTools(cmdTools.UBound).GetHeight + m_buttonMarginBottom
    End If
    
    vOffset = vOffset + m_labelMarginTop
    lblRecording.SetPosition lblRecording.GetLeft, vOffset

End Sub

'Companion function to ReflowToolboxLayout(), above.
Private Sub PositionToolLabel(ByRef targetLabelIndex As Long, ByRef referenceButton As Object, ByRef hOffset As Long, ByRef vOffset As Long)
    
    If m_ShowCategoryLabels Then
        
        Dim heightCalc As Long
        If ttlCategories(targetLabelIndex - 1).Value Then heightCalc = referenceButton.GetHeight Else heightCalc = 0
        
        vOffset = referenceButton.GetTop + heightCalc + m_buttonMarginBottom
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
        targetObject.SetPosition hOffset, vOffset
        
        'If the associated title index is set to TRUE, display the button and calculate a new offset for the next button
        targetObject.Visible = ttlCategories(associatedTitleIndex).Value
        
        If ttlCategories(associatedTitleIndex).Value Then
            hOffset = hOffset + targetObject.GetWidth + m_buttonMarginRight
            If (hOffset + targetObject.GetWidth > m_rightBoundary) Then
                hOffset = m_hOffsetDefaultButton
                vOffset = vOffset + targetObject.GetHeight + m_buttonMarginBottom
            End If
        End If
        
    Next i
    
End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the ToggleToolboxVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    If g_ProgramShuttingDown Then
        ReleaseFormTheming Me
        Set m_MouseEvents = Nothing
        UserPrefs.SetPref_Long "Tools", "LastUsedTool", g_CurrentTool
    Else
        PDDebug.LogAction "WARNING!  toolbar_Toolbox was unloaded prematurely - why??"
        Cancel = True
    End If
End Sub

'When a new tool is selected, we may need to initialize certain values.
Private Sub NewToolSelected()
    
    Select Case g_CurrentTool
        
        'Measure tool
        Case ND_MEASURE
            Tools_Measure.InitializeMeasureTool
            Tools_Measure.ResetPoints True
        
        'Crop tool
        Case ND_CROP
            If Selections.SelectionsAllowed(False) Then Process "Remove selection", createUndo:=UNDO_Selection
            Tools_Crop.InitializeCropTool
            
        'Selection tools
        Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
        
            'See if a selection is already active on the image
            If SelectionsAllowed(False) Then
                
                'A selection is active on this image.  We need to determine if the current selection
                ' shape matches the current selection tool.  If it does, we need to synchronize all
                ' UI elements on the toolpanel to match the active selection.
                If (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
                    
                    'The existing selection type matches the activated selection tool.
                    ' Activate the transformation tools (if any) for this tool, and note
                    ' that this call will also synchronize any UI settings to the active
                    ' selection (e.g. rectangular selections need the position/size/aspect ratio
                    ' controls synchronized to the current values).
                    SetUIGroupState PDUI_SelectionTransforms, PDImages.GetActiveImage.MainSelection.IsTransformable
                    
                'A selection is already active, and it doesn't match the just-activated selection tool.
                Else
                    
                    'Squash composite selections down to a single raster selection.  This frees a lot of resources
                    ' and improves selection performance, and if the user switches away from the active selection
                    ' tool it's basically a sign that they're done transforming that active selection.
                    PDImages.GetActiveImage.MainSelection.SquashCompositeToRaster
                    
                    'Release any locked properties (e.g. locked aspect ratio)
                    PDImages.GetActiveImage.MainSelection.UnlockProperty pdsl_Width
                    PDImages.GetActiveImage.MainSelection.UnlockProperty pdsl_Height
                    PDImages.GetActiveImage.MainSelection.UnlockProperty pdsl_AspectRatio
                    
                End If
            
            'A selection is not active; no further action is required
            End If
        
        Case PAINT_PENCIL
            toolpanel_Pencil.SyncAllPencilSettingsToUI
            
        Case PAINT_SOFTBRUSH
            toolpanel_Paintbrush.SyncAllPaintbrushSettingsToUI
            
        Case PAINT_ERASER
            toolpanel_Eraser.SyncAllPaintbrushSettingsToUI
        
        Case PAINT_CLONE
            toolpanel_Clone.SyncAllPaintbrushSettingsToUI
        
        Case PAINT_FILL
            toolpanel_Fill.SyncAllFillSettingsToUI
            
        Case PAINT_GRADIENT
            toolpanel_Gradient.SyncAllGradientSettingsToUI
            
    End Select
    
    'Because tools may do some custom rendering atop the image canvas, now is a good time to redraw the canvas.
    ' (Note that we can use a relatively late pipeline stage, as only tool-specific overlays need to be redrawn.)
    If PDImages.IsImageActive() Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
                
    'Perform additional per-image initializations, as needed
    Tools.InitializeToolsDependentOnImage
    
    'Finally, free any resources tied to the old tool
    Select Case g_PreviousTool
    
        Case PAINT_SOFTBRUSH
            Tools_Paint.ReduceMemoryIfPossible
            
        Case PAINT_CLONE
            Tools_Clone.ReduceMemoryIfPossible
            
        Case PAINT_FILL
            Tools_Fill.ReduceMemoryIfPossible
    
    End Select
    
    'With all tool settings initialized, set focus to the canvas.  (Because the previous
    ' tool panel was unloaded, focus can be unpredictable if left up to the system.)
    If (Not g_WindowManager Is Nothing) Then FormMain.MainCanvas(0).SetFocusToCanvasView
    
End Sub

'External functions can use this to request the selection of a new tool (for example, Select All uses this to set the
' rectangular tool selector as the current tool)
Public Sub SelectNewTool(ByVal newToolID As PDTools, Optional ByVal flashToolButton As Boolean = False, Optional ByVal setEvenIfAlreadySet As Boolean = False)
    
    Dim okToProceed As Boolean
    okToProceed = setEvenIfAlreadySet
    If (Not okToProceed) Then okToProceed = (newToolID <> g_CurrentTool)
    
    If okToProceed Then
        g_PreviousTool = g_CurrentTool
        g_CurrentTool = newToolID
        ResetToolButtonStates flashToolButton
    End If
    
End Sub

'When a new tool button is selected, we need to raise all the others and display the proper options box
Public Sub ResetToolButtonStates(Optional ByVal flashCurrentButton As Boolean = False)
        
    m_InsideReflowCode = True
        
    'Start by depressing the selected button and raising all unselected ones
    Dim catID As Long
    For catID = 0 To cmdTools.Count - 1
        If (catID = g_CurrentTool) Then
            If (Not cmdTools(catID).Value) Then cmdTools(catID).Value = True
            If flashCurrentButton Then cmdTools(catID).FlashButton
        Else
            If cmdTools(catID).Value Then cmdTools(catID).Value = False
        End If
    Next catID
    
    Dim i As Long
    
    'If our panel tracker doesn't exist, create it now
    If (m_NumOfPanels = 0) Then
        m_NumOfPanels = NUM_OF_TOOL_PANELS
        ReDim m_Panels(0 To m_NumOfPanels - 1) As ToolPanelTracker
    End If
    
    'Next, we need to display the correct tool options panel.  There is no set pattern to this;
    ' some tools share panels, but show/hide certain controls as necessary.  Other tools require
    ' their own unique panel.
    
    'I've tried to strike a balance between "as few panels as possible" without going overboard.
    Select Case g_CurrentTool
        
        'Move/size tool
        Case NAV_MOVE
            Load toolpanel_MoveSize
            toolpanel_MoveSize.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_MoveSize
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_MoveSize.hWnd
            
        'Color picker tool
        Case COLOR_PICKER
            Load toolpanel_ColorPicker
            toolpanel_ColorPicker.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_ColorPicker
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_ColorPicker.hWnd
            
        'Measure tool
        Case ND_MEASURE
            Load toolpanel_Measure
            toolpanel_Measure.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Measure
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_Measure.hWnd
        
        'Crop tool
        Case ND_CROP
            Load toolpanel_Crop
            toolpanel_Crop.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Crop
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_Crop.hWnd
        
        'Rectangular, Elliptical, Line selections
        Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
            Load toolpanel_Selections
            toolpanel_Selections.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Selections
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_Selections.hWnd
            
        'Vector tools
        Case TEXT_BASIC
            Load toolpanel_TextBasic
            toolpanel_TextBasic.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Text
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_TextBasic.hWnd
            
        Case TEXT_ADVANCED
            Load toolpanel_TextAdvanced
            toolpanel_TextAdvanced.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Typography
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_TextAdvanced.hWnd
        
        'Paint tools
        Case PAINT_PENCIL
            Load toolpanel_Pencil
            toolpanel_Pencil.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Pencil
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_Pencil.hWnd
            
        Case PAINT_SOFTBRUSH
            Load toolpanel_Paintbrush
            toolpanel_Paintbrush.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Paintbrush
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_Paintbrush.hWnd
            
        Case PAINT_ERASER
            Load toolpanel_Eraser
            toolpanel_Eraser.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Eraser
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_Eraser.hWnd
        
        Case PAINT_CLONE
            Load toolpanel_Clone
            toolpanel_Clone.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Clone
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_Clone.hWnd
            
        Case PAINT_FILL
            Load toolpanel_Fill
            toolpanel_Fill.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Fill
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_Fill.hWnd
            
        Case PAINT_GRADIENT
            Load toolpanel_Gradient
            toolpanel_Gradient.UpdateAgainstCurrentTheme
            m_ActiveToolPanel = TP_Gradient
            m_Panels(m_ActiveToolPanel).PanelHWnd = toolpanel_Gradient.hWnd
        
        'If a tool does not require an extra settings panel, set the active panel to -1.  This will hide all panels.
        Case Else
            m_ActiveToolPanel = TP_None
            
    End Select
    
    'Notify the parent options panel of the new tool panel.  It will handle window synchronization between the two.
    If (m_ActiveToolPanel = TP_None) Then
        toolbar_Options.NotifyChildPanelHWnd 0
    Else
        toolbar_Options.NotifyChildPanelHWnd m_Panels(m_ActiveToolPanel).PanelHWnd
        m_Panels(m_ActiveToolPanel).PanelWasLoaded = True
    End If
    
    'If a selection tool is active, we also need activate a specific subpanel.  (All selection tools share the same
    ' parent window, but they only activate subportions of it based on tool features.)
    Dim activeSelectionSubpanel As Long
    If Tools.IsSelectionToolActive Then
    
        activeSelectionSubpanel = SelectionUI.GetSelectionSubPanelFromCurrentTool()
        
        For i = 0 To toolpanel_Selections.ctlGroupSelectionSubcontainer.Count - 1
            toolpanel_Selections.ctlGroupSelectionSubcontainer(i).Visible = (i = activeSelectionSubpanel)
        Next i
        
        'When switching tools, we also unlock all locked selection attributes.
        If Selections.SelectionsAllowed(False) Then
            PDImages.GetActiveImage.MainSelection.UnlockProperty pdsl_Width
            PDImages.GetActiveImage.MainSelection.UnlockProperty pdsl_Height
            PDImages.GetActiveImage.MainSelection.UnlockProperty pdsl_AspectRatio
        End If
        
    End If
    
    'Next, some tools display information about the current layer.  Synchronize that information before proceeding,
    ' so that the option panel's information is correct as soon as the window appears.
    Tools.SyncToolOptionsUIToCurrentLayer
    
    'If the current tool is a selection tool, make sure the selection area box (interior/exterior/border) is enabled properly
    If Tools.IsSelectionToolActive Then
        toolpanel_Selections.UpdateSelectionPanelLayout
    
    'Otherwise, squash any composite selections down to a single mask.  (This frees up significant resources.)
    Else
        If Selections.SelectionsAllowed(False) Then
            If PDImages.GetActiveImage.IsSelectionActive Then PDImages.GetActiveImage.MainSelection.SquashCompositeToRaster
        End If
    End If
    
    'Next, we can automatically hide the options toolbox for certain tools (because they have no options).  This is a
    ' nice courtesy, as it frees up space on the main canvas area if the current tool has no adjustable options.
    ' (Note that we can skip the check if the main form is not yet loaded.)
    If FormMain.Visible Then
    
        Select Case g_CurrentTool
            
            'Hand and zoom tools do not provide additional options
            Case NAV_DRAG, NAV_ZOOM
                Toolboxes.SetToolboxVisibility PDT_TopToolbox, False
                
            'All other tools expose options, so display the toolbox (unless the user has disabled the window completely)
            Case Else
                Toolboxes.SetToolboxVisibilityByPreference PDT_TopToolbox
                
        End Select
        
        FormMain.UpdateMainLayout
        
    End If
    
    'Next, we want to display the current tool options panel, while hiding all inactive ones.
    ' (This must be handled carefully, or we risk accidentally enabling unloaded panels,
    '  which we don't want as toolpanels are quite resource-heavy.)
    g_WindowManager.DeactivateToolPanel m_CurrentToolPanelHwnd
    
    'To prevent flicker, we handle this in two passes.
    
    'First, activate the new window.
    If (m_NumOfPanels <> 0) Then
    
        For i = 0 To m_NumOfPanels - 1
            
            'If this is the active panel, display it
            If (i = m_ActiveToolPanel) Then
                g_WindowManager.ActivateToolPanel m_Panels(i).PanelHWnd, toolbar_Options.hWnd
                m_CurrentToolPanelHwnd = m_Panels(i).PanelHWnd
                Exit For
            End If
            
        Next i
        
        'Next, forcibly hide all other panels
        For i = 0 To m_NumOfPanels - 1
            
            If (i <> m_ActiveToolPanel) Then
                    
                'Hide the (now inactive) panel
                If (m_Panels(i).PanelHWnd <> 0) Then
                    g_WindowManager.SetVisibilityByHWnd m_Panels(i).PanelHWnd, False
                    m_Panels(i).PanelHWnd = 0
                End If
                
            End If
            
        Next i
        
    End If
    
    NewToolSelected
    
    m_InsideReflowCode = False
        
End Sub

Private Sub cmdTools_Click(Index As Integer, ByVal Shift As ShiftConstants)
    
    'Update the previous and current tool entries
    If cmdTools(Index).Value Then
        g_PreviousTool = g_CurrentTool
        g_CurrentTool = Index
    End If
    
    'Update the tool options area to match the newly selected tool
    If (Not m_InsideReflowCode) And cmdTools(Index).Value Then
        
        'If the user is dragging the mouse in from the right, and the toolbox has been shrunk from its default setting, the class cursor
        ' for forms may get stuck on the west/east "resize" cursor.  To avoid this, reset it after any button click.
        m_MouseEvents.SetCursor_System IDC_ARROW
        
        'Repaint all tool buttons to reflect the new selection
        PDDebug.LogAction "Selected tool: " & Tools.GetNameOfTool(g_CurrentTool)
        ResetToolButtonStates
    
    End If
    
End Sub

Private Sub ttlCategories_Click(Index As Integer, ByVal newState As Boolean)
    ReflowToolboxLayout
End Sub

'Used to change the visibility of category labels.  When disabled, the button layout is reflowed into a continuous
' stream of buttons.  When enabled, buttons are sorted by category.
Public Sub ToggleToolCategoryLabels(Optional ByVal newSetting As PD_BOOL = PD_BOOL_AUTO)
    
    If (newSetting = PD_BOOL_AUTO) Then
        m_ShowCategoryLabels = (Not m_ShowCategoryLabels)
    ElseIf (newSetting = PD_BOOL_FALSE) Then
        m_ShowCategoryLabels = False
    Else
        m_ShowCategoryLabels = True
    End If
    
    FormMain.MnuWindowToolbox(2).Checked = m_ShowCategoryLabels
    UserPrefs.SetPref_Boolean "Core", "Show Toolbox Category Labels", m_ShowCategoryLabels
    
    'Reflow the interface
    ReflowToolboxLayout
    
End Sub

'Used to change the display size of toolbox buttons.  newSize is expected on the range [0, 2] for small, medium, large
Public Sub UpdateButtonSize(ByVal newSize As PD_ToolboxButtonSize, Optional ByVal suppressRedraw As Boolean = False)
    
    'Export the updated size to file
    If (Not suppressRedraw) Then UserPrefs.SetPref_Long "Core", "Toolbox Button Size", newSize
    
    'Update our internal size metrics to match
    m_ButtonSize = newSize
    If (m_ButtonSize = tbs_Small) Then
        m_ButtonWidth = FixDPI(BTS_WIDTH_SMALL)
        m_ButtonHeight = FixDPI(BTS_HEIGHT_SMALL)
    ElseIf (m_ButtonSize = tbs_Medium) Then
        m_ButtonWidth = FixDPI(BTS_WIDTH_MEDIUM)
        m_ButtonHeight = FixDPI(BTS_HEIGHT_MEDIUM)
    Else
        m_ButtonWidth = FixDPI(BTS_WIDTH_LARGE)
        m_ButtonHeight = FixDPI(BTS_HEIGHT_LARGE)
    End If
    
    'Update all buttons to match
    Dim i As Long
    
    For i = 0 To cmdFile.UBound
        If (cmdFile(i).GetWidth <> m_ButtonWidth) Or (cmdFile(i).GetHeight <> m_ButtonHeight) Then
            cmdFile(i).SetPositionAndSize cmdFile(i).Left, cmdFile(i).Top, m_ButtonWidth, m_ButtonHeight
        End If
    Next i
    
    For i = 0 To cmdTools.UBound
        If (cmdTools(i).GetWidth <> m_ButtonWidth) Or (cmdTools(i).GetHeight <> m_ButtonHeight) Then
            cmdTools(i).SetPositionAndSize cmdTools(i).Left, cmdTools(i).Top, m_ButtonWidth, m_ButtonHeight
        End If
    Next i
    
    'Convert the newSize value to a menu index
    newSize = newSize + 4

    'Apply a checkbox to the matching menu item
    For i = 4 To 6
        FormMain.MnuWindowToolbox(i).Checked = (i = newSize)
    Next i
    
    'TODO: notify the new toolbox manager of this change
    'Toolboxes.FillDefaultToolboxValues??
    'g_WindowManager.UpdateMinimumDimensions Me.hWnd, m_ButtonWidth
    
    'Reflow the interface as requested
    If Me.Visible Then UpdateAgainstCurrentTheme
    If (Not suppressRedraw) Then ReflowToolboxLayout
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'From the current button size, determine what size we want our various image resources
    Dim buttonImageSize As Long
    If (m_ButtonSize = tbs_Small) Then
        buttonImageSize = BTS_IMG_SMALL
    ElseIf (m_ButtonSize = tbs_Medium) Then
        buttonImageSize = BTS_IMG_MEDIUM
    ElseIf (m_ButtonSize = tbs_Large) Then
        buttonImageSize = BTS_IMG_LARGE
    End If
    
    buttonImageSize = Interface.FixDPI(buttonImageSize)
    
    'Initialize file tool button images
    cmdFile(FILE_NEW).AssignImage "file_new", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
    cmdFile(FILE_OPEN).AssignImage "file_open", Nothing, buttonImageSize, buttonImageSize
    cmdFile(FILE_CLOSE).AssignImage "file_close", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
    cmdFile(FILE_SAVE).AssignImage "file_save", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
    cmdFile(FILE_SAVEAS_LAYERS).AssignImage "file_savedup", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
    cmdFile(FILE_SAVEAS_FLAT).AssignImage "file_saveas", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
    
    cmdFile(FILE_UNDO).AssignImage "edit_undo", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
    cmdFile(FILE_REDO).AssignImage "edit_redo", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
    
    'Initialize canvas tool button images
    cmdTools(NAV_DRAG).AssignImage "nd_hand", Nothing, buttonImageSize, buttonImageSize
    cmdTools(NAV_ZOOM).AssignImage "zoom_default", Nothing, buttonImageSize, buttonImageSize
    cmdTools(NAV_MOVE).AssignImage "nd_move", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
    cmdTools(COLOR_PICKER).AssignImage "color_picker", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
    cmdTools(ND_MEASURE).AssignImage "nd_measure", Nothing, buttonImageSize, buttonImageSize, resampleAlgorithm:=GP_IM_NearestNeighbor
    cmdTools(ND_CROP).AssignImage "image_crop", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=rf_CatmullRom
    
    cmdTools(SELECT_RECT).AssignImage "select_rect", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=rf_Box
    cmdTools(SELECT_CIRC).AssignImage "select_circle", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=rf_CatmullRom
    cmdTools(SELECT_POLYGON).AssignImage "select_polygon", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=rf_CatmullRom
    cmdTools(SELECT_LASSO).AssignImage "select_lasso", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=rf_CatmullRom
    cmdTools(SELECT_WAND).AssignImage "select_wand", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
    
    cmdTools(TEXT_BASIC).AssignImage "text_basic", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
    cmdTools(TEXT_ADVANCED).AssignImage "text_fancy", Nothing, buttonImageSize, buttonImageSize
    
    cmdTools(PAINT_PENCIL).AssignImage "paint_pencil", Nothing, buttonImageSize, buttonImageSize
    cmdTools(PAINT_SOFTBRUSH).AssignImage "paint_softbrush", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
    cmdTools(PAINT_ERASER).AssignImage "paint_erase", Nothing, buttonImageSize, buttonImageSize
    cmdTools(PAINT_CLONE).AssignImage "clone_stamp", Nothing, buttonImageSize, buttonImageSize
    cmdTools(PAINT_FILL).AssignImage "paint_fill", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
    cmdTools(PAINT_GRADIENT).AssignImage "nd_gradient", Nothing, buttonImageSize, buttonImageSize, usePDResamplerInstead:=rf_CatmullRom
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me
    
    'Tooltips must be manually re-assigned according to the current language.  This is a necessary evil, if the user switches
    ' between two non-English languages at run-time.
    
    'Title bars first
    ttlCategories(0).AssignTooltip "File tools: create, load, and save image files"
    ttlCategories(1).AssignTooltip "Undo/Redo tools: revert changes made to an image or layer"
    ttlCategories(2).AssignTooltip "Layout tools: explore, zoom, move or measure an image or layer"
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
    
    If UserPrefs.GetPref_Long("Saving", "Overwrite Or Copy", 0) = 0 Then
        cmdFile(FILE_SAVE).AssignTooltip "WARNING: this will overwrite the current image file.  To save to a different file, use the ""Save As"" button.", "Save image in current format"
    Else
        cmdFile(FILE_SAVE).AssignTooltip "You have specified ""safe"" save mode, which means that each save will create a new file with an auto-incremented filename.", "Save image in current format"
    End If
    
    cmdFile(FILE_SAVEAS_LAYERS).AssignTooltip "Use this to quickly save a lossless copy of the current image.  The lossless copy will be saved in PDI format, in the image's current folder, using the current filename (plus an auto-incremented number, as necessary).", "Save lossless copy"
    cmdFile(FILE_SAVEAS_FLAT).AssignTooltip "The Save As command always raises a dialog, so you can specify a new file name, folder, and/or image format for the current image.", "Save As (export to new format or filename)"
        
    'Layout tool buttons are next
    Dim shortcutText As String, hotkeyText As String
    shortcutText = g_Language.TranslateMessage("Hand (click-and-drag image scrolling)")
    If Hotkeys.GetHotkeyText_FromAction("tool_hand", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(NAV_DRAG).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Zoom")
    If Hotkeys.GetHotkeyText_FromAction("tool_zoom", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(NAV_ZOOM).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Move and resize image layers")
    If Hotkeys.GetHotkeyText_FromAction("tool_move", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(NAV_MOVE).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Select colors from the image")
    If Hotkeys.GetHotkeyText_FromAction("tool_colorselect", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(COLOR_PICKER).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Measure angles and distances")
    If Hotkeys.GetHotkeyText_FromAction("tool_colorselect", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(ND_MEASURE).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Crop")
    If Hotkeys.GetHotkeyText_FromAction("tool_crop", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(ND_CROP).AssignTooltip shortcutText
    
    '...then selections...
    shortcutText = g_Language.TranslateMessage("Rectangular Selection")
    If Hotkeys.GetHotkeyText_FromAction("tool_select_rect", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(SELECT_RECT).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Elliptical (Oval) Selection")
    If Hotkeys.GetHotkeyText_FromAction("tool_select_rect", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(SELECT_CIRC).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Polygon Selection")
    If Hotkeys.GetHotkeyText_FromAction("tool_select_polygon", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(SELECT_POLYGON).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Lasso (Freehand) Selection")
    If Hotkeys.GetHotkeyText_FromAction("tool_select_polygon", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(SELECT_LASSO).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Magic Wand Selection")
    If Hotkeys.GetHotkeyText_FromAction("tool_select_wand", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(SELECT_WAND).AssignTooltip shortcutText
    
    '...then vector tools...
    shortcutText = g_Language.TranslateMessage("Basic Text")
    If Hotkeys.GetHotkeyText_FromAction("tool_text_basic", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(TEXT_BASIC).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Advanced Text")
    If Hotkeys.GetHotkeyText_FromAction("tool_text_basic", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(TEXT_ADVANCED).AssignTooltip shortcutText
    
    '...then paint tools...
    shortcutText = g_Language.TranslateMessage("Pencil (hard-tipped brush)")
    If Hotkeys.GetHotkeyText_FromAction("tool_pencil", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(PAINT_PENCIL).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Paintbrush (soft-tipped brush)")
    If Hotkeys.GetHotkeyText_FromAction("tool_paintbrush", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(PAINT_SOFTBRUSH).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Eraser")
    If Hotkeys.GetHotkeyText_FromAction("tool_erase", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(PAINT_ERASER).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Clone stamp")
    If Hotkeys.GetHotkeyText_FromAction("tool_clone", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(PAINT_CLONE).AssignTooltip shortcutText
    
    shortcutText = g_Language.TranslateMessage("Paint bucket (fill with color)")
    If Hotkeys.GetHotkeyText_FromAction("tool_paintbucket", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(PAINT_FILL).AssignTooltip shortcutText
    shortcutText = g_Language.TranslateMessage("Gradient")
    If Hotkeys.GetHotkeyText_FromAction("tool_gradient", hotkeyText) Then shortcutText = shortcutText & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", hotkeyText)
    cmdTools(PAINT_GRADIENT).AssignTooltip shortcutText
    
    'And finally, tool names and their corresponding action strings.  (These are supplied to the
    ' hotkey manager, which is why we use slightly different organization, combining some tools that
    ' are designed to share hotkeys.)
    ReDim m_ToolActions(0) As PD_ToolboxAction
    m_numToolActions = 0
    
    AddToolboxAction g_Language.TranslateMessage("Hand tool"), "tool_hand"
    AddToolboxAction g_Language.TranslateMessage("Zoom tool"), "tool_zoom"
    AddToolboxAction g_Language.TranslateMessage("Move tool"), "tool_move"
    AddToolboxAction g_Language.TranslateMessage("Color selector tool") & ", " & g_Language.TranslateMessage("Measure tool"), "tool_colorselect"
    AddToolboxAction g_Language.TranslateMessage("Crop tool"), "tool_crop"
    AddToolboxAction g_Language.TranslateMessage("Rectangle selection tool") & ", " & g_Language.TranslateMessage("Ellipse selection tool"), "tool_select_rect"
    AddToolboxAction g_Language.TranslateMessage("Polygon selection tool") & ", " & g_Language.TranslateMessage("Lasso selection tool"), "tool_select_polygon"
    AddToolboxAction g_Language.TranslateMessage("Magic wand selection tool"), "tool_select_wand"
    AddToolboxAction g_Language.TranslateMessage("Basic text tool") & ", " & g_Language.TranslateMessage("Advanced text tool"), "tool_text_basic"
    AddToolboxAction g_Language.TranslateMessage("Pencil tool"), "tool_pencil"
    AddToolboxAction g_Language.TranslateMessage("Paintbrush tool"), "tool_paintbrush"
    AddToolboxAction g_Language.TranslateMessage("Erase tool"), "tool_erase"
    AddToolboxAction g_Language.TranslateMessage("Clone stamp tool"), "tool_clone"
    AddToolboxAction g_Language.TranslateMessage("Paint bucket tool"), "tool_paintbucket"
    AddToolboxAction g_Language.TranslateMessage("Gradient tool"), "tool_gradient"
    AddToolboxAction g_Language.TranslateMessage("Search tool"), "tool_search"
    
    'Tool modifiers; UI setting changes only!
    AddToolboxAction g_Language.TranslateMessage("Decrease brush size"), "tool_active_sizedown"
    AddToolboxAction g_Language.TranslateMessage("Increase brush size"), "tool_active_sizeup"
    AddToolboxAction g_Language.TranslateMessage("Decrease brush hardness"), "tool_active_hardnessdown"
    AddToolboxAction g_Language.TranslateMessage("Increase brush hardness"), "tool_active_hardnessup"
    AddToolboxAction g_Language.TranslateMessage("Toggle brush cursor"), "tool_active_togglecursor"
    
End Sub

Private Sub AddToolboxAction(ByRef translatedName As String, ByRef toolAction As String)
    
    If (m_numToolActions = 0) Then
        Const INIT_TOOL_ACTIONS As Long = 16
        ReDim m_ToolActions(0 To INIT_TOOL_ACTIONS - 1) As PD_ToolboxAction
    End If
    
    If (m_numToolActions > UBound(m_ToolActions)) Then ReDim Preserve m_ToolActions(0 To m_numToolActions * 2 - 1) As PD_ToolboxAction
    
    With m_ToolActions(m_numToolActions)
        .ta_ToolName = translatedName
        .ta_ToolAction = toolAction
    End With
    
    m_numToolActions = m_numToolActions + 1
    
End Sub

Public Sub GetListOfToolNamesAndActions(ByRef dstNames As pdStringStack, ByRef dstActions As pdStringStack)
    
    Set dstNames = New pdStringStack
    Set dstActions = New pdStringStack
    
    If (m_numToolActions > 0) Then
        Dim i As Long
        For i = 0 To m_numToolActions - 1
            dstNames.AddString m_ToolActions(i).ta_ToolName
            dstActions.AddString m_ToolActions(i).ta_ToolAction
        Next i
    End If
    
End Sub

'You *must* call this function before shutdown.  This function will forcibly free cached toolbox windows.
Public Sub FreeAllToolpanels()
    
    'If a flyout panel is open on the current toolbar, close it
    UserControls.HideOpenFlyouts 0&
    
    'The active toolpanel (if one exists) has had its window bits manually modified so that we can
    ' embed it atop the parent tool options window.  Make certain those window bits are reset before
    ' we attempt to unload the panel using built-in VB keywords (because VB will crash if it
    ' encounters unexpected window bits, especially WS_CHILD).
    If (Not g_WindowManager Is Nothing) And (m_CurrentToolPanelHwnd <> 0) Then g_WindowManager.DeactivateToolPanel m_CurrentToolPanelHwnd
    m_CurrentToolPanelHwnd = 0
    
    'Make sure our internal toolbox collection actually exists before attempting to iterate it
    If (m_NumOfPanels = 0) Then Exit Sub
    
    'Free any toolboxes that were loaded this session
    Dim i As PD_ToolPanels
    For i = 0 To NUM_OF_TOOL_PANELS - 1
        
        'If we loaded this panel during this session, unload it manually now
        If m_Panels(i).PanelWasLoaded Then
        
            Select Case i
                Case TP_MoveSize
                    Unload toolpanel_MoveSize
                    Set toolpanel_MoveSize = Nothing
                Case TP_ColorPicker
                    Unload toolpanel_ColorPicker
                    Set toolpanel_ColorPicker = Nothing
                Case TP_Measure
                    Unload toolpanel_Measure
                    Set toolpanel_Measure = Nothing
                Case TP_Selections
                    Unload toolpanel_Selections
                    Set toolpanel_Selections = Nothing
                Case TP_Text
                    Unload toolpanel_TextBasic
                    Set toolpanel_TextBasic = Nothing
                Case TP_Typography
                    Unload toolpanel_TextAdvanced
                    Set toolpanel_TextAdvanced = Nothing
                Case TP_Pencil
                    Unload toolpanel_Pencil
                    Set toolpanel_Pencil = Nothing
                Case TP_Paintbrush
                    Unload toolpanel_Paintbrush
                    Set toolpanel_Paintbrush = Nothing
                Case TP_Eraser
                    Unload toolpanel_Eraser
                    Set toolpanel_Eraser = Nothing
                Case TP_Clone
                    Unload toolpanel_Clone
                    Set toolpanel_Clone = Nothing
                Case TP_Fill
                    Unload toolpanel_Fill
                    Set toolpanel_Fill = Nothing
                Case TP_Gradient
                    Unload toolpanel_Gradient
                    Set toolpanel_Gradient = Nothing
            End Select
            
            m_Panels(i).PanelWasLoaded = False
            
        End If
        
    Next i
    
End Sub

Private Sub ttlCategories_CustomDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Loading.LoadFromDragDrop Data, Effect, Button, Shift
End Sub

Private Sub ttlCategories_CustomDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub
