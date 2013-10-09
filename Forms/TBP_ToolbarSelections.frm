VERSION 5.00
Begin VB.Form toolbar_Selections 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selections"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3045
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
   ScaleHeight     =   521
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6135
      Index           =   0
      Left            =   0
      ScaleHeight     =   409
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   2970
      Begin VB.ComboBox cmbSelRender 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "TBP_ToolbarSelections.frx":0000
         Left            =   180
         List            =   "TBP_ToolbarSelections.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   540
         Width           =   2685
      End
      Begin VB.ComboBox cmbSelType 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "TBP_ToolbarSelections.frx":0004
         Left            =   120
         List            =   "TBP_ToolbarSelections.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "This option controls the selection's area.  You can switch between the three settings without losing the current selection."
         Top             =   4320
         Width           =   2685
      End
      Begin VB.ComboBox cmbSelSmoothing 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "TBP_ToolbarSelections.frx":0008
         Left            =   120
         List            =   "TBP_ToolbarSelections.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Use this option to change the way selections blend with their surroundings."
         Top             =   3000
         Width           =   2685
      End
      Begin PhotoDemon.sliderTextCombo sltCornerRounding 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   5640
         Width           =   3000
         _ExtentX        =   5318
         _ExtentY        =   873
         Max             =   10000
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
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
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
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   1320
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
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
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
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
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   3
         Left            =   1560
         TabIndex        =   12
         Top             =   2160
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
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
      Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   4710
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5318
         _ExtentY        =   873
         Min             =   1
         Max             =   10000
         Value           =   1
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
      Begin PhotoDemon.sliderTextCombo sltSelectionFeathering 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   0
         TabIndex        =   14
         Top             =   3390
         Width           =   3000
         _ExtentX        =   5318
         _ExtentY        =   873
         Max             =   100
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
      Begin PhotoDemon.sliderTextCombo sltSelectionLineWidth 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   0
         TabIndex        =   15
         Top             =   5640
         Width           =   3000
         _ExtentX        =   5318
         _ExtentY        =   873
         Min             =   1
         Max             =   10000
         Value           =   10
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
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "selection size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1380
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "selection position"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1830
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "visual style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "corner rounding"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   5280
         Width           =   1710
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "selection type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   3960
         Width           =   1440
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "smoothing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "Select (Ellipse)"
      Height          =   600
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Elliptical (Oval) Selection tool"
      Top             =   495
      Width           =   900
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "Select (Rect)"
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Rectangular Selection tool"
      Top             =   495
      Width           =   900
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "Select (Line)"
      Height          =   600
      Index           =   2
      Left            =   2040
      TabIndex        =   0
      ToolTipText     =   "Line Selection tool"
      Top             =   495
      Width           =   900
   End
   Begin VB.Label lblTools 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "selection tools"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   285
      Left            =   165
      TabIndex        =   3
      Top             =   120
      Width           =   1500
   End
   Begin VB.Line lineMain 
      BorderColor     =   &H80000002&
      Index           =   1
      X1              =   8
      X2              =   195
      Y1              =   94
      Y2              =   94
   End
End
Attribute VB_Name = "toolbar_Selections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Selections Toolbar
'Copyright ©2012-2013 by Tanner Helland
'Created: 03/October/13
'Last updated: 03/October/13
'Last update: initial build
'
'This form was initially integrated into the main MDI form.  In fall 2013, PhotoDemon left behind the MDI model,
' and all toolbars were moved to their own forms.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Used to toggle the command button state of the toolbox buttons
Private Const BM_SETSTATE = &HF3
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'Used to render images onto the tool buttons at run-time
' NOTE: TOOLBOX IMAGES WILL NOT APPEAR IN THE IDE.  YOU MUST COMPILE FIRST.
Private cImgCtl As clsControlImage

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'Because VB doesn't provide a good equivalent for the "Show" event, and because the Load event is too early, we must
' manually track the first activation, and load any last-used settings then.
Private hasBeenActivated As Boolean


'Upon first activation, restore all last-used settings
Private Sub Form_Activate()

    If Not hasBeenActivated Then
        hasBeenActivated = True
        
        'Load any last-used settings for this form
        Set lastUsedSettings = New pdLastUsedSettings
        lastUsedSettings.setParentForm Me
        'lastUsedSettings.loadAllControlValues
        
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    'lastUsedSettings.saveAllControlValues

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
        g_CurrentTool = SELECT_RECT
    End If
    cmdTools_Click CInt(g_CurrentTool)
        
    Dim i As Long
    For i = 0 To tudSel.Count - 1
        tudSel(i) = 0
    Next i

End Sub

'When the selection type is changed, update the corresponding preference and redraw all selections
Private Sub cmbSelRender_Click(Index As Integer)
            
    If NumOfWindows > 0 Then
    
        Dim i As Long
        For i = 0 To NumOfImagesLoaded
            If (Not pdImages(i) Is Nothing) Then
                If pdImages(i).IsActive And pdImages(i).selectionActive Then RenderViewport pdImages(i).containingForm
            End If
        Next i
    
    End If
    
End Sub

'Change selection smoothing (e.g. none, antialiased, fully feathered)
Private Sub cmbSelSmoothing_Click(Index As Integer)
    
    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(CurrentImage).mainSelection.setSmoothingType cmbSelSmoothing(Index).ListIndex
        pdImages(CurrentImage).mainSelection.setFeatheringRadius sltSelectionFeathering.Value
        RenderViewport pdImages(CurrentImage).containingForm
    End If
    
End Sub

'Change selection type (e.g. interior, exterior, bordered)
Private Sub cmbSelType_Click(Index As Integer)

    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(CurrentImage).mainSelection.setSelectionType cmbSelType(Index).ListIndex
        pdImages(CurrentImage).mainSelection.setBorderSize sltSelectionBorder.Value
        RenderViewport pdImages(CurrentImage).containingForm
    End If
    
End Sub

'External functions can use this to request the selection of a new tool (for example, Select All uses this to set the
' rectangular tool selector as the current tool)
Public Sub selectNewTool(ByVal newToolID As Long)
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = newToolID
    resetToolButtonStates
End Sub

Private Sub cmdTools_Click(Index As Integer)
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = Index
    resetToolButtonStates
End Sub

'Private Sub cmdTools_LostFocus(Index As Integer)
    'g_CurrentTool = Index
    'resetToolButtonStates
'End Sub

Private Sub cmdTools_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    g_CurrentTool = Index
    resetToolButtonStates
End Sub

'When a new tool button is selected, we need to raise all the others and display the proper options box
Public Sub resetToolButtonStates()
    
    'Start by depressing the selected button and raising all unselected ones
    Dim i As Long
    For i = 0 To cmdTools.Count - 1
        SendMessageA cmdTools(i).hWnd, BM_SETSTATE, False, 0
    Next i
    SendMessageA cmdTools(g_CurrentTool).hWnd, BM_SETSTATE, True, 0
    
    'Next, we need to display the correct tool options panel.  There is no set pattern to this; some tools share
    ' panels, but show/hide certain controls as necessary.  Other tools require their own unique panel.  I've tried
    ' to strike a balance between "as few panels as possible" without going overboard.
    Dim activeToolPanel As Long
    
    Select Case g_CurrentTool
        
        'Rectangular, Elliptical selections
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
            activeToolPanel = 0
        
        Case Else
        
    End Select
    
    'If tools share the same panel, they may need to show or hide a few additional controls.  (For example,
    ' "corner rounding", which is needed for rectangular selections but not elliptical ones, despite the two
    ' sharing the same tool panel.)  Do this before showing or hiding the tool panel.
    Select Case g_CurrentTool
    
        'For rectangular selections, show the rounded corners option
        Case SELECT_RECT
            toolbar_Selections.lblSelection(5).Visible = True
            toolbar_Selections.sltCornerRounding.Visible = True
            toolbar_Selections.sltSelectionLineWidth.Visible = False
            
        'For elliptical selections, hide the rounded corners option
        Case SELECT_CIRC
            toolbar_Selections.lblSelection(5).Visible = False
            toolbar_Selections.sltCornerRounding.Visible = False
            toolbar_Selections.sltSelectionLineWidth.Visible = False
            
        'Line selections also show the rounded corners slider, though they repurpose it for line width
        Case SELECT_LINE
            toolbar_Selections.lblSelection(5).Visible = True
            toolbar_Selections.sltCornerRounding.Visible = False
            toolbar_Selections.sltSelectionLineWidth.Visible = True
        
    End Select
    
    'Even if tools share the same panel, they may name controls differently, or use different max/min values.
    ' Check for this, and apply new text and max/min settings as necessary.
    Select Case g_CurrentTool
    
        'Rectangular and elliptical selections use rectangular bounding boxes and potential corner rounding
        Case SELECT_RECT, SELECT_CIRC
            lblSelection(1).Caption = g_Language.TranslateMessage("selection position")
            lblSelection(2).Caption = g_Language.TranslateMessage("selection size")
            lblSelection(5).Caption = g_Language.TranslateMessage("corner rounding")
            'If (g_PreviousTool <> SELECT_RECT) And (g_PreviousTool <> SELECT_CIRC) Then
                'If selectionsAllowed And (Not g_UndoRedoActive) Then Process "Remove selection", , , 2, g_PreviousTool
                'If g_CurrentTool = SELECT_RECT Then sltCornerRounding.Value = 0
            'End If
            
        'Line selections use two points, and the corner rounding slider gets repurposed as line width.
        Case SELECT_LINE
            'If selectionsAllowed And (Not g_UndoRedoActive) Then Process "Remove selection", , , 2, g_PreviousTool
            lblSelection(1).Caption = g_Language.TranslateMessage("first point (x, y)")
            lblSelection(2).Caption = g_Language.TranslateMessage("second point (x, y)")
            lblSelection(5).Caption = g_Language.TranslateMessage("line width")
            
    End Select
    
    'Display the current tool options panel, while hiding all inactive ones
    For i = 0 To picTools.Count - 1
        If i = activeToolPanel Then
            If Not picTools(i).Visible Then
                picTools(i).Visible = True
                setArrowCursorToObject picTools(i)
            End If
        Else
            If picTools(i).Visible Then picTools(i).Visible = False
        End If
    Next i
    
    newToolSelected
        
End Sub

'When a new tool is selected, we may need to initialize certain values
Private Sub newToolSelected()
    
    Select Case g_CurrentTool
    
        'Rectangular, elliptical selections
        Case SELECT_RECT
                
            'If a similar selection is already active, change its shape to match the current tool, then redraw it
            If selectionsAllowed(True) And (Not g_UndoRedoActive) Then
                If (g_PreviousTool = SELECT_CIRC) And (pdImages(CurrentImage).mainSelection.getSelectionShape = sCircle) Then
                    pdImages(CurrentImage).mainSelection.setSelectionShape g_CurrentTool
                    RenderViewport pdImages(CurrentImage).containingForm
                Else
                    If pdImages(CurrentImage).mainSelection.getSelectionShape = sRectangle Then
                        metaToggle tSelectionTransform, True
                    Else
                        metaToggle tSelectionTransform, False
                    End If
                End If
            End If
            
        Case SELECT_CIRC
        
            'If a similar selection is already active, change its shape to match the current tool, then redraw it
            If selectionsAllowed(True) And (Not g_UndoRedoActive) Then
                If (g_PreviousTool = SELECT_RECT) And (pdImages(CurrentImage).mainSelection.getSelectionShape = sRectangle) Then
                    pdImages(CurrentImage).mainSelection.setSelectionShape g_CurrentTool
                    RenderViewport pdImages(CurrentImage).containingForm
                Else
                    If pdImages(CurrentImage).mainSelection.getSelectionShape = sCircle Then
                        metaToggle tSelectionTransform, True
                    Else
                        metaToggle tSelectionTransform, False
                    End If
                End If
            End If
            
        'Line selections
        Case SELECT_LINE
        
            'Deactivate the position text boxes - those shouldn't be accessible unless a line selection is presently active
            If selectionsAllowed(True) Then
                If pdImages(CurrentImage).mainSelection.getSelectionShape = sLine Then
                    metaToggle tSelectionTransform, True
                Else
                    metaToggle tSelectionTransform, False
                End If
            Else
                metaToggle tSelectionTransform, False
            End If
            
        Case Else
        
    End Select
    
End Sub

Private Sub Form_Load()

    'Render images to the toolbox command buttons
    Dim i As Long
    
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    ' (NOTE: because the icons require manifest theming, they will not appear in the IDE.)
    Set cImgCtl = New clsControlImage
    If g_IsProgramCompiled Then
        
        'Remove all tool button captions
        For i = 0 To cmdTools.Count - 1
            cmdTools(i).Caption = ""
        Next i
        
        With cImgCtl
            
            'Load the tool images (in PNG format) from the resource file
            .LoadImageFromStream cmdTools(0).hWnd, LoadResData("T_SELRECT", "CUSTOM"), 22, 22
            .LoadImageFromStream cmdTools(1).hWnd, LoadResData("T_SELCIRCLE", "CUSTOM"), 22, 22
            .LoadImageFromStream cmdTools(2).hWnd, LoadResData("T_SELLINE", "CUSTOM"), 22, 22
            
            'Center-align the images in their respective buttons
            For i = 0 To cmdTools.Count - 1
                .SetMargins cmdTools(i).hWnd, 0
                .Align(cmdTools(i).hWnd) = Icon_Center
                
                'On XP, the tool command button images aren't aligned properly until the buttons are hovered.  No one
                ' knows why.  We can imitate a hover with a click - do so now.
                If Not g_IsVistaOrLater Then cmdTools_Click CInt(i)
            Next i
            
        End With
    End If
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    g_WindowManager.unregisterForm Me
End Sub

Private Sub sltCornerRounding_Change()
    If selectionsAllowed(True) Then
        pdImages(CurrentImage).mainSelection.setRoundedCornerAmount sltCornerRounding.Value
        RenderViewport pdImages(CurrentImage).containingForm
    End If
End Sub

Private Sub sltSelectionBorder_Change()
    If selectionsAllowed(False) Then
        pdImages(CurrentImage).mainSelection.setBorderSize sltSelectionBorder.Value
        RenderViewport pdImages(CurrentImage).containingForm
    End If
End Sub

Private Sub sltSelectionFeathering_Change()
    If selectionsAllowed(False) Then
        pdImages(CurrentImage).mainSelection.setFeatheringRadius sltSelectionFeathering.Value
        RenderViewport pdImages(CurrentImage).containingForm
    End If
End Sub

Private Sub sltSelectionLineWidth_Change()
    If selectionsAllowed(True) Then
        pdImages(CurrentImage).mainSelection.setSelectionLineWidth sltSelectionLineWidth.Value
        RenderViewport pdImages(CurrentImage).containingForm
    End If
End Sub

Private Function selectionsAllowed(ByVal transformableMatters As Boolean) As Boolean
    If NumOfWindows > 0 Then
        If pdImages(CurrentImage).selectionActive And (Not pdImages(CurrentImage).mainSelection Is Nothing) And (Not pdImages(CurrentImage).mainSelection.rejectRefreshRequests) Then
            
            If transformableMatters Then
                If pdImages(CurrentImage).mainSelection.isTransformable Then
                    selectionsAllowed = True
                Else
                    selectionsAllowed = False
                End If
            Else
                selectionsAllowed = True
            End If
            
        Else
            selectionsAllowed = False
        End If
    Else
        selectionsAllowed = False
    End If
End Function

'When certain selection settings are enabled or disabled, corresponding controls are shown or hidden.  To keep the
' panel concise and clean, we move other controls up or down depending on what controls are visible.
Private Sub updateSelectionPanelLayout()

    'Display the feathering slider as necessary
    If cmbSelSmoothing(0).ListIndex = sFullyFeathered Then
        sltSelectionFeathering.Visible = True
        lblSelection(4).Top = sltSelectionFeathering.Top + fixDPI(38)
    Else
        sltSelectionFeathering.Visible = False
        lblSelection(4).Top = cmbSelSmoothing(0).Top + fixDPI(34)
    End If
    cmbSelType(0).Top = lblSelection(4).Top + fixDPI(24)
    sltSelectionBorder.Top = cmbSelType(0).Top + fixDPI(26)

    'Display the border slider as necessary
    If cmbSelType(0).ListIndex = sBorder Then
        sltSelectionBorder.Visible = True
        lblSelection(5).Top = sltSelectionBorder.Top + fixDPI(38)
    Else
        sltSelectionBorder.Visible = False
        lblSelection(5).Top = cmbSelType(0).Top + fixDPI(34)
    End If
    sltCornerRounding.Top = lblSelection(5).Top + fixDPI(24)
    sltSelectionLineWidth.Top = lblSelection(5).Top + fixDPI(24)

End Sub

'When the selection text boxes are updated, change the scrollbars to match
Private Sub tudSel_Change(Index As Integer)
    updateSelectionsValuesViaText
End Sub

Private Sub updateSelectionsValuesViaText()
    If selectionsAllowed(True) Then
        If Not pdImages(CurrentImage).mainSelection.rejectRefreshRequests Then
            pdImages(CurrentImage).mainSelection.updateViaTextBox
            RenderViewport pdImages(CurrentImage).containingForm
        End If
    End If
End Sub
