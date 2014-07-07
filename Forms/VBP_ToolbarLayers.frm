VERSION 5.00
Begin VB.Form toolbar_Layers 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Layers"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3735
   FillStyle       =   0  'Solid
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
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboBlendMode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   675
      Width           =   2655
   End
   Begin VB.VScrollBar vsLayer 
      Height          =   4905
      LargeChange     =   32
      Left            =   3360
      Max             =   100
      TabIndex        =   7
      Top             =   1320
      Width           =   285
   End
   Begin VB.PictureBox picLayerButtons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   3
      Top             =   6750
      Width           =   3735
      Begin PhotoDemon.jcbutton cmdLayerAction 
         Height          =   480
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   847
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_ToolbarLayers.frx":0000
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
         ColorScheme     =   3
      End
      Begin PhotoDemon.jcbutton cmdLayerAction 
         Height          =   480
         Index           =   3
         Left            =   2640
         TabIndex        =   5
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   847
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_ToolbarLayers.frx":0D52
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
         ColorScheme     =   3
      End
      Begin PhotoDemon.jcbutton cmdLayerAction 
         Height          =   480
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   847
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_ToolbarLayers.frx":1AA4
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
         ColorScheme     =   3
      End
      Begin PhotoDemon.jcbutton cmdLayerAction 
         Height          =   480
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   847
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_ToolbarLayers.frx":27F6
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
         ColorScheme     =   3
      End
   End
   Begin VB.PictureBox picLayers 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   327
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   215
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3255
      Begin VB.TextBox txtLayerName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin PhotoDemon.sliderTextCombo sltLayerOpacity 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   100
      Value           =   100
   End
   Begin VB.Label lblLayerSettings 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "blend:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   240
      Index           =   0
      Left            =   375
      TabIndex        =   9
      Top             =   720
      Width           =   540
   End
   Begin VB.Line lnSeparator 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   8
      X2              =   240
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Label lblLayerSettings 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "opacity:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   675
   End
End
Attribute VB_Name = "toolbar_Layers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Layers Toolbar
'Copyright ©2013-2014 by Tanner Helland
'Created: 25/March/14
'Last updated: 01/June/14
'Last update: implement drag/drop layer reordering
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'A collection of all currently active layer thumbnails.  It is dynamically resized as layers are added/removed.
' For performance reasons, we cache it locally, and only update it as necessary.  Also, layers are referenced by their
' canonical ID rather than their layer order - important, as order can obviously change!
Private Type thumbEntry
    thumbDIB As pdDIB
    canonicalLayerID As Long
End Type

Private layerThumbnails() As thumbEntry
Private numOfThumbnails As Long

'Until I settle on final thumb width/height values, I've declared them as variables.
Private thumbWidth As Long, thumbHeight As Long

'I don't want thumbnails to fill the full size of their individual blocks, so a border of this many pixels is automatically
' applied to each side of the thumbnail.  (Like all other interface elements, it is dynamically modified for DPI as necessary.)
Private Const thumbBorder As Long = 3

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'An outside class provides access to mousewheel events for scrolling the layer view
Private WithEvents cMouseEvents As pdInput
Attribute cMouseEvents.VB_VarHelpID = -1

'A separate mouse handler helps with dynamically resizing the layer toolbox
Private WithEvents cMouseEventsForm As pdInput
Attribute cMouseEventsForm.VB_VarHelpID = -1

'Height of each layer content block.  Note that this is effectively a "magic number", in pixels, representing the
' height of each layer block in the layer selection UI.  This number will be dynamically resized per the current
' screen DPI by the "redrawLayerList" and "renderLayerBlock" functions.
Private Const BLOCKHEIGHT As Long = 48

'The distance (in pixels at 96 dpi) between clickable buttons in the "show on hover" layer block menu
Private Const DIST_BETWEEN_HOVER_BUTTONS As Long = 12

'Internal DIB (and measurements) for the custom layer list interface
Private bufferDIB As pdDIB
Private m_BufferWidth As Long, m_BufferHeight As Long

'A font object, used for rendering layer names, and its color (set by Form_Load, and which will eventually be themed).
Private layerNameFont As pdFont, layerNameColor As Long

'The currently hovered layer entry.  (Note that the currently *selected* layer entry is retrieved from the active
' pdImage object, rather than stored locally.)
Private curLayerHover As Long

'Layer buttons are more easily referenced by this enum rather than their actual indices
Private Enum LAYER_BUTTON_ID
    LYR_BTN_ADD = 0
    LYR_BTN_DELETE = 1
    LYR_BTN_MOVE_UP = 2
    LYR_BTN_MOVE_DOWN = 3
End Enum

#If False Then
    Private Const LYR_BTN_ADD = 0, LYR_BTN_DELETE = 1, LYR_BTN_MOVE_UP = 2, LYR_BTN_MOVE_DOWN = 3
#End If

'Sometimes we need to make changes that will raise redraw-causing events.  Set this variable to TRUE if you want
' such functions to ignore their automatic redrawing.
Private disableRedraws As Boolean

'Extra interface images are loaded as resources at run-time
Private img_EyeOpen As pdDIB, img_EyeClosed As pdDIB
Private img_MergeUp As pdDIB, img_MergeDown As pdDIB, img_MergeUpDisabled As pdDIB, img_MergeDownDisabled As pdDIB
Private img_Duplicate As pdDIB

'Some UI elements are dynamically rendered onto the layer box.  To simplify hit detection, their RECTs are stored
' at render-time, which allows the mouse actions to easily check hits regardless of layer box position.
Private m_VisibilityRect As RECT, m_NameRect As RECT
Private m_MergeUpRect As RECT, m_MergeDownRect As RECT
Private m_DuplicateRect As RECT

'While in OLE drag/drop mode (e.g. dragging files from Explorer), ignore any mouse actions on the main layer box
Private m_InOLEDragDropMode As Boolean

'While in our own custom layer box drag/drop mode (e.g. rearranging layers), this will be set to TRUE.
' Also, the layer-to-be-moved is tracked, as is the initial layer index (which is required for processing the final
' action, e.g. the one that triggers Undo/Redo creation).
Private m_LayerRearrangingMode As Boolean, m_LayerIndexToRearrange As Long, m_InitialLayerIndex As Long

'When we are responsible for this window resizing (because the user is resizing our window manually), we set this to TRUE.
' This variable is then checked before requesting additional redraws during our resize event.
Private weAreResponsibleForResize As Boolean

'When the mouse is over the layer list, this will be set to TRUE
Private m_MouseOverLayerBox As Boolean

'External functions can force a full redraw by calling this sub.  (This is necessary whenever layers are added, deleted,
' re-ordered, etc.)
Public Sub forceRedraw(Optional ByVal refreshThumbnailCache As Boolean = True)
    
    If refreshThumbnailCache Then cacheLayerThumbnails
    
    'Sync opacity, blend mode, and other controls to the currently active layer
    disableRedraws = True
    If (g_OpenImageCount > 0) Then
        If (Not pdImages(g_CurrentImage).getActiveLayer Is Nothing) Then
            
            'Synchronize the opacity scroll bar to the active layer
            sltLayerOpacity.Value = pdImages(g_CurrentImage).getActiveLayer.getLayerOpacity
            
            'Synchronize the blend mode to the active layer
            cboBlendMode.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getLayerBlendMode
            
        End If
    End If
    
    disableRedraws = False
    
    'resizeLayerUI already calls all the proper redraw functions for us, so simply link it here
    resizeLayerUI
    
    'Determine which buttons need to be activated.
    checkButtonEnablement
    
End Sub

'Whenever a layer is activated, we must re-determine which buttons the user has access to.  Move up/down are disabled for
' entries at either end, and the last layer of an image cannot be deleted.
Private Sub checkButtonEnablement()

    'Make sure at least one image has been loaded
    If (Not pdImages(g_CurrentImage) Is Nothing) And (g_OpenImageCount > 0) Then

        'Add layer is always allowed
        cmdLayerAction(LYR_BTN_ADD).Enabled = True
        
        'Merge down is only allowed for layer indexes > 0
        If pdImages(g_CurrentImage).getActiveLayerIndex = 0 Then
            cmdLayerAction(LYR_BTN_MOVE_DOWN).Enabled = False
        Else
            cmdLayerAction(LYR_BTN_MOVE_DOWN).Enabled = True
        End If
        
        'Merge up is only allowed for layer indexes < NUM_OF_LAYERS
        If pdImages(g_CurrentImage).getActiveLayerIndex < pdImages(g_CurrentImage).getNumOfLayers - 1 Then
            cmdLayerAction(LYR_BTN_MOVE_UP).Enabled = True
        Else
            cmdLayerAction(LYR_BTN_MOVE_UP).Enabled = False
        End If
        
        'Delete layer is only allowed if there are multiple layers present
        If pdImages(g_CurrentImage).getNumOfLayers > 1 Then
            cmdLayerAction(LYR_BTN_DELETE).Enabled = True
        Else
            cmdLayerAction(LYR_BTN_DELETE).Enabled = False
        End If
    
    'If no images are loaded, disable all layer action buttons
    Else
    
        Dim i As Long
        For i = cmdLayerAction.lBound To cmdLayerAction.UBound
            cmdLayerAction(i).Enabled = False
        Next i
        
    End If
    
End Sub

'Change the blend mode of the active layer
Private Sub cboBlendMode_Click()

    'By default, changing the drop-down will automatically update the blend mode of the selected layer, and the main viewport
    ' will be redrawn.  When changing the blend mode programmatically, set disableRedraws to TRUE to prevent cylical redraws.
    If disableRedraws Then Exit Sub

    If g_OpenImageCount > 0 Then
    
        If Not pdImages(g_CurrentImage).getActiveLayer Is Nothing Then
        
            pdImages(g_CurrentImage).getActiveLayer.setLayerBlendMode cboBlendMode.ListIndex
            ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        End If
    
    End If

End Sub

'Layer action buttons - move layers up/down, delete layers, etc.
Private Sub cmdLayerAction_Click(Index As Integer)

    Dim copyOfCurLayerID As Long
    copyOfCurLayerID = pdImages(g_CurrentImage).getActiveLayerID

    Select Case Index
    
        Case LYR_BTN_ADD
            Process "Add new layer", True
        
        Case LYR_BTN_DELETE
            Process "Delete layer", False, pdImages(g_CurrentImage).getActiveLayerIndex, UNDO_IMAGE
        
        Case LYR_BTN_MOVE_UP
            Process "Raise layer", False, pdImages(g_CurrentImage).getActiveLayerIndex, UNDO_IMAGEHEADER
        
        Case LYR_BTN_MOVE_DOWN
            Process "Lower layer", False, pdImages(g_CurrentImage).getActiveLayerIndex, UNDO_IMAGEHEADER
            
    End Select
    
End Sub

'Clicks on the layer box raise all kinds of fun events, depending on where they occur
Private Sub cMouseEvents_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    Dim clickedLayer As Long
    clickedLayer = getLayerAtPosition(x, y)
    
    If clickedLayer >= 0 Then
        
        If (Not pdImages(g_CurrentImage) Is Nothing) And (Button = pdLeftButton) Then
            
            'Check the clicked position against a series of rects, each one representing a unique interaction.
            
            'Has the user clicked a visibility rectangle?
            If isPointInRect(x, y, m_VisibilityRect) Then
                
                Layer_Handler.setLayerVisibilityByIndex clickedLayer, Not pdImages(g_CurrentImage).getLayerByIndex(clickedLayer).getLayerVisibility, True
            
            'Duplicate rectangle?
            ElseIf isPointInRect(x, y, m_DuplicateRect) Then
            
                Process "Duplicate Layer", False, Str(clickedLayer), UNDO_IMAGE
            
            'Merge down rectangle?
            ElseIf isPointInRect(x, y, m_MergeDownRect) Then
            
                If Layer_Handler.isLayerAllowedToMergeAdjacent(clickedLayer, True) >= 0 Then
                    Process "Merge layer down", False, Str(clickedLayer), UNDO_IMAGE
                End If
            
            'Merge up rectangle?
            ElseIf isPointInRect(x, y, m_MergeUpRect) Then
            
                If Layer_Handler.isLayerAllowedToMergeAdjacent(clickedLayer, False) >= 0 Then
                    Process "Merge layer up", False, Str(clickedLayer), UNDO_IMAGE
                End If
            
            'The user has not clicked any item of interest.  Assume that they want to make the clicked layer
            ' the active layer.
            Else
                Layer_Handler.setActiveLayerByIndex clickedLayer, False
                RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
            End If
            
            'Redraw the layer box to represent any changes from this interaction.
            ' NOTE: this is not currently necessary, as all interactions will automatically force a redraw on their own.
            'redrawLayerBox
            
        End If
        
    End If
    
End Sub

'Double-clicks on the layer box raise "layer title edit mode", if the mouse is within a layer's title area
Private Sub cMouseEvents_DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub

    If isPointInRect(x, y, m_NameRect) And (Button = pdLeftButton) Then
    
        'Move the text layer box into position
        txtLayerName.Move m_NameRect.Left, m_NameRect.Top, m_NameRect.Right - m_NameRect.Left, m_NameRect.Bottom - m_NameRect.Top
        txtLayerName.Visible = True
        
        'Disable hotkeys until editing is finished
        FormMain.ctlAccelerator.Enabled = False
        
        'Fill the text box with the current layer name, and select it
        txtLayerName.Text = pdImages(g_CurrentImage).getLayerByIndex(getLayerAtPosition(x, y)).getLayerName
        AutoSelectText txtLayerName
    
    Else
    
        'Hide the text box if it isn't already
        txtLayerName.Visible = False
    
    End If

End Sub

'Arrow keys have been pressed (or number pad arrow keys)
Private Sub cMouseEvents_KeyDownArrows(ByVal Shift As ShiftConstants, ByVal upArrow As Boolean, ByVal rightArrow As Boolean, ByVal downArrow As Boolean, ByVal leftArrow As Boolean)
    
    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    'Ignore arrow keys unless an image has been loaded
    If Not pdImages(g_CurrentImage) Is Nothing Then
    
        'Up key activates the next layer upward
        If upArrow And (pdImages(g_CurrentImage).getActiveLayerIndex < pdImages(g_CurrentImage).getNumOfLayers - 1) Then
            Layer_Handler.setActiveLayerByIndex pdImages(g_CurrentImage).getActiveLayerIndex + 1, True
        End If
        
        'Down key activates the next layer downward
        If downArrow And pdImages(g_CurrentImage).getActiveLayerIndex > 0 Then
            Layer_Handler.setActiveLayerByIndex pdImages(g_CurrentImage).getActiveLayerIndex - 1, True
        End If
        
        'Right key increases active layer opacity
        If rightArrow And (pdImages(g_CurrentImage).getActiveLayer.getLayerVisibility) Then
            sltLayerOpacity.Value = sltLayerOpacity.Value + 10
            ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        End If
        
        'Left key decreases active layer opacity
        If leftArrow And (pdImages(g_CurrentImage).getActiveLayer.getLayerVisibility) Then
            sltLayerOpacity.Value = sltLayerOpacity.Value - 10
            ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        End If
        
    End If

End Sub

'An edit key (http://en.wikipedia.org/wiki/Template:Keyboard_keys) has been pressed.  Note that PD strives to provide the same hotkeys for
' both this layer toolbox and the Move/Resize tool, so mirror any changes here to the pdCanvas user control as well!
Private Sub cMouseEvents_KeyDownEdits(ByVal Shift As ShiftConstants, ByVal kReturn As Boolean, ByVal kEnter As Boolean, ByVal kSpaceBar As Boolean, ByVal kBackspace As Boolean, ByVal kInsert As Boolean, ByVal kDelete As Boolean, ByVal kTab As Boolean, ByVal kEscape As Boolean)

    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    'Ignore key presses unless an image has been loaded
    If Not pdImages(g_CurrentImage) Is Nothing Then
    
        'Delete key: delete the active layer (if allowed)
        If kDelete And pdImages(g_CurrentImage).getNumOfLayers > 1 Then
            Process "Delete layer", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGE
        End If
        
        'Insert: raise Add New Layer dialog
        If kInsert Then
            Process "Add new layer", True
            
            'Recapture focus
            picLayers.SetFocus
        End If
        
        'Tab and Shift+Tab: move through layer stack
        If kTab Then
            
            'Retrieve the active layer index
            Dim curLayerIndex As Long
            curLayerIndex = pdImages(g_CurrentImage).getActiveLayerIndex
            
            'Advance the layer index according to the Shift modifier
            If (Shift And vbShiftMask) <> 0 Then
                curLayerIndex = curLayerIndex + 1
            Else
                curLayerIndex = curLayerIndex - 1
            End If
            
            If curLayerIndex < 0 Then curLayerIndex = pdImages(g_CurrentImage).getNumOfLayers - 1
            If curLayerIndex > pdImages(g_CurrentImage).getNumOfLayers - 1 Then curLayerIndex = 0
            
            'Activate the new layer
            pdImages(g_CurrentImage).setActiveLayerByIndex curLayerIndex
            
            'Redraw the viewport and interface to match
            RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
            syncInterfaceToCurrentImage
            
        End If
        
        'Space bar: toggle active layer visibility
        If kSpaceBar Then
            pdImages(g_CurrentImage).getActiveLayer.setLayerVisibility (Not pdImages(g_CurrentImage).getActiveLayer.getLayerVisibility)
            ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
            syncInterfaceToCurrentImage
        End If
    
    End If

End Sub

'MouseDown is used to process our own custom layer drag/drop reordering
Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    'Retrieve the layer under this position
    Dim clickedLayer As Long
    clickedLayer = getLayerAtPosition(x, y)
    
    'Don't proceed unless the user has the mouse over a valid layer
    If (clickedLayer >= 0) And (Not pdImages(g_CurrentImage) Is Nothing) Then
        
        'If the image is a multilayer image, and they're using the left mouse button, initiate drag/drop layer reordering
        If (pdImages(g_CurrentImage).getNumOfLayers > 1) And (Button = pdLeftButton) Then
        
            'Enter layer rearranging mode
            m_LayerRearrangingMode = True
            
            'Note the layer being rearranged
            m_LayerIndexToRearrange = clickedLayer
            m_InitialLayerIndex = m_LayerIndexToRearrange
        
        End If
        
    End If

End Sub

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseOverLayerBox = True
End Sub

'Mouse has left the layer box
Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    m_MouseOverLayerBox = False

    'Note that no layer is currently hovered
    curLayerHover = -1
    
    'Redraw the layer box, which no longer has anything hovered
    redrawLayerBox

End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    'Only display the hand cursor if the cursor is over a layer
    If getLayerAtPosition(x, y) <> -1 Then
        cMouseEvents.setSystemCursor IDC_HAND
    Else
        cMouseEvents.setSystemCursor IDC_ARROW
    End If
    
    'Don't process further MouseMove events if no images are loaded
    If (g_OpenImageCount = 0) Or (pdImages(g_CurrentImage) Is Nothing) Then Exit Sub
    
    'Process any important interactions first.  If a live interaction is taking place (such as drag/drop layer reordering),
    ' other MouseMove events will be suspended until the drag/drop is completed.
    
    'Check for drag/drop reordering
    If m_LayerRearrangingMode Then
    
        'The user is in the middle of a drag/drop reorder.  Give them a live update!
        
        'Retrieve the layer under this position
        Dim layerIndexUnderMouse As Long
        layerIndexUnderMouse = getLayerAtPosition(x, y, True)
                
        'Ask the parent pdImage to move the layer for us
        If pdImages(g_CurrentImage).moveLayerToArbitraryIndex(m_LayerIndexToRearrange, layerIndexUnderMouse) Then
        
            'Note that the layer currently being moved has changed
            m_LayerIndexToRearrange = layerIndexUnderMouse
            
            'Keep the current layer as the active one
            setActiveLayerByIndex layerIndexUnderMouse, False
            
            'Redraw the layer box, and note that thumbnails need to be re-cached
            toolbar_Layers.forceRedraw True
            
            'Redraw the viewport
            ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        End If
        
    End If
    
    'If a layer other than the active one is being hovered, highlight that box
    If curLayerHover <> getLayerAtPosition(x, y) Then
        curLayerHover = getLayerAtPosition(x, y)
        redrawLayerBox
    End If
    
    'Update the tooltip contingent on the mouse position.
    Dim toolString As String
    
    'Mouse is over a visibility toggle
    If isPointInRect(x, y, m_VisibilityRect) Then
        
        'Fast mouse movements can cause this event to trigger, even when no layer is hovered.
        ' As such, we need to make sure we won't be attempting to access a bad layer index.
        If curLayerHover >= 0 Then
            If pdImages(g_CurrentImage).getLayerByIndex(curLayerHover).getLayerVisibility Then
                toolString = g_Language.TranslateMessage("Click to hide this layer.")
            Else
                toolString = g_Language.TranslateMessage("Click to show this layer.")
            End If
        End If
        
    'Mouse is over Duplicate
    ElseIf isPointInRect(x, y, m_DuplicateRect) Then
    
        If curLayerHover >= 0 Then
            toolString = g_Language.TranslateMessage("Click to duplicate this layer.")
        End If
    
    'Mouse is over Merge Down
    ElseIf isPointInRect(x, y, m_MergeDownRect) Then
    
        If curLayerHover >= 0 Then
            If Layer_Handler.isLayerAllowedToMergeAdjacent(curLayerHover, True) >= 0 Then
                toolString = g_Language.TranslateMessage("Click to merge this layer with the layer beneath it.")
            Else
                toolString = g_Language.TranslateMessage("This layer can't merge down, because there are no visible layers beneath it.")
            End If
        End If
            
    'Mouse is over Merge Up
    ElseIf isPointInRect(x, y, m_MergeUpRect) Then
    
        If curLayerHover >= 0 Then
            If Layer_Handler.isLayerAllowedToMergeAdjacent(curLayerHover, False) >= 0 Then
                toolString = g_Language.TranslateMessage("Click to merge this layer with the layer above it.")
            Else
                toolString = g_Language.TranslateMessage("This layer can't merge up, because there are no visible layers above it.")
            End If
        End If
            
    'The user has not clicked any item of interest.  Assume that they want to make the clicked layer
    ' the active layer.
    Else
        
        'The tooltip is irrelevant if the current layer is already active
        If pdImages(g_CurrentImage).getActiveLayerIndex <> getLayerAtPosition(x, y) Then
            toolString = g_Language.TranslateMessage("Click to make this the active layer.")
        Else
            toolString = g_Language.TranslateMessage("This is the currently active layer.")
        End If
        
    End If
    
    'Only update the tooltip if it differs from the current one.  (This prevents horrific flickering.)
    If StrComp(m_ToolTip.ToolText(picLayers), toolString, vbTextCompare) <> 0 Then m_ToolTip.ToolText(picLayers) = toolString
    
End Sub

'MouseUp
Private Sub cMouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)

    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    'Retrieve the layer under this position
    Dim layerIndexUnderMouse As Long
    layerIndexUnderMouse = getLayerAtPosition(x, y, True)
    
    'Don't proceed further unless an image has been loaded, and the user is not just clicking the layer box
    If (Not pdImages(g_CurrentImage) Is Nothing) And (Not ClickEventAlsoFiring) Then
        
        'If we're in drag/drop mode, and the left mouse button is pressed, terminate drag/drop layer reordering
        If m_LayerRearrangingMode And (Button = pdLeftButton) Then
        
            'Exit layer rearranging mode
            m_LayerRearrangingMode = False
            
            'Ask the parent pdImage to move the layer for us; the MouseMove event has probably taken care of this already.
            ' In that case, this function will return FALSE and we don't have to do anything extra.
            If pdImages(g_CurrentImage).moveLayerToArbitraryIndex(m_LayerIndexToRearrange, layerIndexUnderMouse) Then
    
                'Keep the current layer as the active one
                setActiveLayerByIndex layerIndexUnderMouse, False
                
                'Redraw the layer box, and note that thumbnails need to be re-cached
                toolbar_Layers.forceRedraw True
                
                'Redraw the viewport
                ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                
            End If
            
            'If the new position differs from the layer's original position, call a dummy Processor call, which will create
            ' an Undo/Redo entry at this point.
            If m_InitialLayerIndex <> layerIndexUnderMouse Then Process "Rearrange layers", False, "", UNDO_IMAGEHEADER
        
        End If
        
    End If
    
    'If we haven't already, exit layer rearranging mode
    m_LayerRearrangingMode = False

End Sub

Private Sub cMouseEvents_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

    'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
    If vsLayer.Visible Then
  
        If scrollAmount < 0 Then
            
            If vsLayer.Value + vsLayer.LargeChange > vsLayer.Max Then
                vsLayer.Value = vsLayer.Max
            Else
                vsLayer.Value = vsLayer.Value + vsLayer.LargeChange
            End If
            
            curLayerHover = getLayerAtPosition(x, y)
            redrawLayerBox
        
        ElseIf scrollAmount > 0 Then
            
            If vsLayer.Value - vsLayer.LargeChange < vsLayer.Min Then
                vsLayer.Value = vsLayer.Min
            Else
                vsLayer.Value = vsLayer.Value - vsLayer.LargeChange
            End If
            
            curLayerHover = getLayerAtPosition(x, y)
            redrawLayerBox
            
        End If
        
    End If

End Sub

'Forward any key events from the form to the layer box handler
Private Sub cMouseEventsForm_KeyDownArrows(ByVal Shift As ShiftConstants, ByVal upArrow As Boolean, ByVal rightArrow As Boolean, ByVal downArrow As Boolean, ByVal leftArrow As Boolean)
    Call cMouseEvents_KeyDownArrows(Shift, upArrow, rightArrow, downArrow, leftArrow)
End Sub

Private Sub cMouseEventsForm_KeyDownEdits(ByVal Shift As ShiftConstants, ByVal kReturn As Boolean, ByVal kEnter As Boolean, ByVal kSpaceBar As Boolean, ByVal kBackspace As Boolean, ByVal kInsert As Boolean, ByVal kDelete As Boolean, ByVal kTab As Boolean, ByVal kEscape As Boolean)
    Call cMouseEvents_KeyDownEdits(Shift, kReturn, kEnter, kSpaceBar, kBackspace, kInsert, kDelete, kTab, kEscape)
End Sub

'Forward mousewheel events to the layer box handler
Private Sub cMouseEventsForm_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    Call cMouseEvents_MouseWheelVertical(Button, Shift, x, y, scrollAmount)
End Sub

Private Sub Form_Load()
    
    'Populate the blend mode box
    cboBlendMode.Clear
    cboBlendMode.AddItem "Normal", 0
    cboBlendMode.AddItem "Darken", 1
    cboBlendMode.AddItem "Multiply", 2
    cboBlendMode.AddItem "Color burn", 3
    cboBlendMode.AddItem "Linear burn", 4
    cboBlendMode.AddItem "Lighten", 5
    cboBlendMode.AddItem "Screen", 6
    cboBlendMode.AddItem "Color dodge", 7
    cboBlendMode.AddItem "Linear dodge", 8
    cboBlendMode.AddItem "Overlay", 9
    cboBlendMode.AddItem "Soft light", 10
    cboBlendMode.AddItem "Hard light", 11
    cboBlendMode.AddItem "Vivid light", 12
    cboBlendMode.AddItem "Linear light", 13
    cboBlendMode.AddItem "Pin light", 14
    cboBlendMode.AddItem "Hard mix", 15
    cboBlendMode.AddItem "Difference", 16
    cboBlendMode.AddItem "Exclusion", 17
    cboBlendMode.AddItem "Subtract", 18
    cboBlendMode.AddItem "Divide", 19
    cboBlendMode.AddItem "Hue", 20
    cboBlendMode.AddItem "Saturation", 21
    cboBlendMode.AddItem "Color", 22
    cboBlendMode.AddItem "Luminosity", 23
    cboBlendMode.ListIndex = 0
    
    'Reset the thumbnail array
    numOfThumbnails = 0
    ReDim layerThumbnails(0 To numOfThumbnails) As thumbEntry

    'Activate the custom tooltip handler
    Set m_ToolTip = New clsToolTip
    m_ToolTip.Create Me
    m_ToolTip.MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
    m_ToolTip.DelayTime(ttDelayShow) = 10000
    m_ToolTip.AddTool picLayers, ""
    
    'Add helpful tooltips to the layer action buttons at the bottom of the toolbox
    With m_ToolTip
        .AddTool cmdLayerAction(0), g_Language.TranslateMessage("Add a blank layer to the image.")
        .AddTool cmdLayerAction(1), g_Language.TranslateMessage("Delete the currently selected layer.")
        .AddTool cmdLayerAction(2), g_Language.TranslateMessage("Move the current layer upward in the layer stack.")
        .AddTool cmdLayerAction(3), g_Language.TranslateMessage("Move the current layer downward in the layer stack.")
    End With
    
    'Theme the form
    makeFormPretty Me
    
    'Enable custom input handling for the layer box
    Set cMouseEvents = New pdInput
    cMouseEvents.addInputTracker picLayers.hWnd, True, True, , True
    cMouseEvents.requestKeyTracking picLayers.hWnd
    cMouseEvents.setKeyTrackers picLayers.hWnd, True, True, True
    m_MouseOverLayerBox = False
    
    'Enable simple mouse handling for the form as well
    Set cMouseEventsForm = New pdInput
    cMouseEventsForm.addInputTracker Me.hWnd, True, , , True
    cMouseEventsForm.requestKeyTracking Me.hWnd
    cMouseEventsForm.setKeyTrackers Me.hWnd, True, True, True
    
    'To prevent the parent form's cursor handler from overriding that of the child, we must manually notify pdInput to
    ' ignore cursor handling in certain situations.
    cMouseEventsForm.setCursorOverrideState True
    
    'No layer has been hovered yet
    curLayerHover = -1
    
    'Rearranging mode is not active
    m_LayerRearrangingMode = False
    
    'Prepare a DIB for rendering the Layer box
    Set bufferDIB = New pdDIB
    resizeLayerUI
    
    'Initialize a custom font object for printing layer names
    layerNameColor = RGB(64, 64, 64)
    
    Set layerNameFont = New pdFont
    With layerNameFont
        .setFontColor layerNameColor
        .setFontBold False
        .setFontSize 10
        .setTextAlignment vbLeftJustify
        .createFontObject
    End With
    
    'Load various interface images from the resource
    initializeUIDib img_EyeOpen, "EYE_OPEN"
    initializeUIDib img_EyeClosed, "EYE_CLOSE"
    initializeUIDib img_Duplicate, "DUPL_LAYER"
    initializeUIDib img_MergeUp, "MERGE_UP"
    initializeUIDib img_MergeDown, "MERGE_DOWN"
    initializeUIDib img_MergeUpDisabled, "MERGE_UP"
    initializeUIDib img_MergeDownDisabled, "MERGE_DOWN"
    
    'If a UI image can be disabled, make a grayscale copy of it in advance
    Filters_Layers.GrayscaleDIB img_MergeUpDisabled, True
    Filters_Layers.GrayscaleDIB img_MergeDownDisabled, True
    
    'Force the blend mode drop-down to display the full list of items.  The SendMessage call is structured as:
    ' 1) hWnd of combo box
    ' 2) Constant for "set number of visible items" request
    ' 3) Number of visible items requested (in our case, the count of the combo box)
    ' 4) (unused)
    ' NOTE: this requires a manifest, so it won't work in the IDE.
    Const CB_SETMINVISIBLE As Long = 339
    SendMessage cboBlendMode.hWnd, 339, CLng(cboBlendMode.ListCount), ByVal 0&
    
End Sub

'Load a UI image from the resource section and into a DIB
Private Sub initializeUIDib(ByRef dstDIB As pdDIB, ByRef resString As String)
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB

    loadResourceToDIB resString, tmpDIB
    
    Set dstDIB = New pdDIB
    
    'If the screen is high DPI, resize all DIBs to match
    If fixDPIFloat(1) > 1 Then
        dstDIB.createBlank fixDPI(tmpDIB.getDIBWidth), fixDPI(tmpDIB.getDIBHeight), tmpDIB.getDIBColorDepth, 0
        GDIPlusResizeDIB dstDIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight, tmpDIB, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, InterpolationModeHighQualityBicubic
    Else
        dstDIB.createFromExistingDIB tmpDIB
    End If
        
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
    If (y > 0) And (y < Me.ScaleHeight) And (x < resizeBorderAllowance) Then
        mouseInResizeTerritory = True
        hitCode = HTLEFT
    End If
    
    'If the left mouse button is down, and the mouse is in resize territory, initiate an API resize event
    If mouseInResizeTerritory Then
    
        'Change the cursor to a resize cursor
        cMouseEventsForm.setSystemCursor IDC_SIZEWE
        
        If (Button = vbLeftButton) Then
            weAreResponsibleForResize = True
            ReleaseCapture
            SendMessage Me.hWnd, WM_NCLBUTTONDOWN, hitCode, ByVal 0&
            
            'Notify the window manager of the change, so it can reflow any neighboring windows
            'g_WindowManager.notifyToolboxResized
            
        End If
        
    Else
        If Not m_MouseOverLayerBox Then cMouseEventsForm.setSystemCursor IDC_ARROW
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    weAreResponsibleForResize = False
End Sub

Private Sub Form_Resize()
    reflowInterface
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
        toggleToolbarVisibility LAYER_TOOLBOX
    End If
    
End Sub

'For performance reasons, PD does all layer box rendering to an internal DIB, which is only flipped to the screen as necessary.
' Whenever the toolbox is resized, we must recreate this DIB.
Private Sub resizeLayerUI()

    'Resize the DIB to be the same size as the Layer UI box
    bufferDIB.createBlank picLayers.ScaleWidth, picLayers.ScaleHeight
    
    'Initialize a few other variables now (for performance reasons)
    m_BufferWidth = picLayers.ScaleWidth
    m_BufferHeight = picLayers.ScaleHeight
    
    'Determine thumbnail height/width
    thumbHeight = fixDPI(BLOCKHEIGHT) - fixDPI(2)
    thumbWidth = thumbHeight
    
    'Redraw the toolbar
    redrawLayerBox
    
End Sub

'Cache all current layer thumbnails.  This is required for things like the user switching to a new image, which requires
' us to wipe the current layer cache and start anew.
Private Sub cacheLayerThumbnails()

    'Do not attempt to cache thumbnails if there are no open images
    If (Not pdImages(g_CurrentImage) Is Nothing) And (g_OpenImageCount > 0) Then
    
        'Make sure the active image has at least one layer.  (This should always be true, but better safe than sorry.)
        If pdImages(g_CurrentImage).getNumOfLayers > 0 Then
    
            'Retrieve the number of layers in the current image and prepare the thumbnail cache
            numOfThumbnails = pdImages(g_CurrentImage).getNumOfLayers
            ReDim layerThumbnails(0 To numOfThumbnails - 1) As thumbEntry
            
            'Only cache thumbnails if the active image has one or more layers
            If numOfThumbnails > 0 Then
            
                Dim i As Long
                For i = 0 To numOfThumbnails - 1
                    
                    'Retrieve a thumbnail and ID for this layer
                    layerThumbnails(i).canonicalLayerID = pdImages(g_CurrentImage).getLayerByIndex(i).getLayerID
                    
                    Set layerThumbnails(i).thumbDIB = New pdDIB
                    pdImages(g_CurrentImage).getLayerByIndex(i).requestThumbnail layerThumbnails(i).thumbDIB, thumbHeight - (fixDPI(thumbBorder) * 2)
                    
                Next i
            
            End If
        
        End If
        
    End If
    
    'See if the vertical scroll bar needs to be displayed
    updateLayerScrollbarVisibility
                
End Sub

'When an action occurs that potentially affects the visibility of the vertical scroll bar (such as resizing the form
' vertically, or adding a new layer to the image), call this function to update the scroll bar visibility as necessary.
Private Sub updateLayerScrollbarVisibility()

    'Determine if the vertical scrollbar needs to be visible or not (because there are so many layers that they overflow the box)
    Dim maxLayerBoxSize As Long
    maxLayerBoxSize = fixDPIFloat(BLOCKHEIGHT) * numOfThumbnails - 1
    
    vsLayer.Value = 0
    If maxLayerBoxSize < picLayers.ScaleHeight Then
        
        'Hide the layer box scroll bar
        vsLayer.Visible = False
        vsLayer.Value = 0
        
        'Extend the layer box to be the full size of the form
        picLayers.Width = (vsLayer.Left + vsLayer.Width) - picLayers.Left
        
    Else
        
        'Show the layer box scroll bar
        vsLayer.Visible = True
        vsLayer.Max = maxLayerBoxSize - picLayers.ScaleHeight
        
        'Shrink the layer box so that it does not cover the vertical scroll bar
        picLayers.Width = (vsLayer.Left - picLayers.Left)
        
    End If

End Sub

'Draw the layer box (from scratch)
Private Sub redrawLayerBox()

    'Determine an offset based on the current scroll bar value
    Dim scrollOffset As Long
    scrollOffset = vsLayer.Value
    
    'Erase the current DIB
    bufferDIB.createBlank m_BufferWidth, m_BufferHeight
    
    'If the image has one or more layers, render them to the list.
    If (Not pdImages(g_CurrentImage) Is Nothing) And (g_OpenImageCount > 0) Then
    
        If pdImages(g_CurrentImage).getNumOfLayers > 0 Then
        
            'Loop through the current layer list, drawing layers as we go
            Dim i As Long
            For i = 0 To pdImages(g_CurrentImage).getNumOfLayers - 1
                renderLayerBlock (pdImages(g_CurrentImage).getNumOfLayers - 1) - i, 0, fixDPI(i * BLOCKHEIGHT) - scrollOffset - fixDPI(2)
            Next i
            
        End If
        
    End If
    
    'Copy the buffer to its container picture box
    BitBlt picLayers.hDC, 0, 0, m_BufferWidth, m_BufferHeight, bufferDIB.getDIBDC, 0, 0, vbSrcCopy
    picLayers.Picture = picLayers.Image
    'picLayers.Refresh

End Sub

'Render an individual "block" for a given layer (including name, thumbnail, and a few button toggles)
Private Sub renderLayerBlock(ByVal blockIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current block if it will be visible
    If ((offsetY + fixDPI(BLOCKHEIGHT)) > 0) And (offsetY < m_BufferHeight) Then
    
        offsetY = offsetY + fixDPI(2)
        
        Dim linePadding As Long
        linePadding = fixDPI(2)
    
        Dim mHeight As Single
        Dim tmpRect As RECT
        Dim hBrush As Long
        
        'For performance reasons, retrieve a reference to the corresponding pdLayer object.  We need to
        ' pull a lot of information from this object as part of rendering this block.
        Dim tmpLayerRef As pdLayer
        Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(blockIndex)
        
        'If this layer is the active layer, draw the background with the system's current selection color
        If tmpLayerRef.getLayerID = pdImages(g_CurrentImage).getActiveLayerID Then
        
            SetRect tmpRect, offsetX, offsetY, m_BufferWidth, offsetY + fixDPI(BLOCKHEIGHT)
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FillRect bufferDIB.getDIBDC, tmpRect, hBrush
            DeleteObject hBrush
            
            'Also, color the fonts with the matching highlighted text color (otherwise they won't be readable)
            layerNameFont.setFontColor ConvertSystemColor(vbHighlightText)
        
        'This layer is not the active layer
        Else
        
            'Render the layer name in a standard, non-highlighted font
            layerNameFont.setFontColor layerNameColor
        
            'If the current layer is mouse-hovered (but not active), render its border with a highlight
            If (blockIndex = curLayerHover) Then
                SetRect tmpRect, offsetX, offsetY, m_BufferWidth, offsetY + fixDPI(BLOCKHEIGHT)
                hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
                FrameRect bufferDIB.getDIBDC, tmpRect, hBrush
                DeleteObject hBrush
            End If
            
        End If
        
        'Object offsets are stored in these values as various elements are drawn to the screen.
        Dim xObjOffset As Long, yObjOffset As Long
        
        'Render the layer thumbnail.  If the layer is not currently visible, render it at 30% opacity.
        xObjOffset = offsetX + fixDPI(thumbBorder)
        yObjOffset = offsetY + fixDPI(thumbBorder)
        If Not (layerThumbnails(blockIndex).thumbDIB Is Nothing) Then
        
            If tmpLayerRef.getLayerVisibility Then
                layerThumbnails(blockIndex).thumbDIB.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
            Else
                layerThumbnails(blockIndex).thumbDIB.alphaBlendToDC bufferDIB.getDIBDC, 76, xObjOffset, yObjOffset
                
                'Also, render a "closed eye" icon in the corner.
                ' NOTE: I'm not sold on this being a good idea.  The icon seems to be clickable, but it isn't!
                'img_EyeClosed.alphaBlendToDC bufferDIB.getDIBDC, 210, xObjOffset + (BLOCKHEIGHT - img_EyeClosed.getDIBWidth) - fixDPI(5), yObjOffset + (BLOCKHEIGHT - img_EyeClosed.getDIBHeight) - fixDPI(6)
                
            End If
            
        End If
        
        'Render the layer name
        Dim drawString As String
        drawString = tmpLayerRef.getLayerName
        
        'If this layer is invisible, mark it as such.
        ' NOTE: not sold on this behavior, but I'm leaving it for a bit to see how it affects workflow.
        If Not tmpLayerRef.getLayerVisibility Then drawString = g_Language.TranslateMessage("(hidden)") & " " & drawString
        
        layerNameFont.attachToDC bufferDIB.getDIBDC
        
        Dim xTextOffset As Long, yTextOffset As Long, xTextWidth As Long, yTextHeight As Long
        xTextOffset = offsetX + thumbWidth + fixDPI(thumbBorder) * 2
        yTextOffset = offsetY + fixDPI(4)
        xTextWidth = m_BufferWidth - xTextOffset - fixDPI(4)
        yTextHeight = layerNameFont.getHeightOfString(drawString)
        layerNameFont.fastRenderTextWithClipping xTextOffset, yTextOffset, xTextWidth, yTextHeight, drawString
        
        'Store the resulting text area in the text rect; if the user clicks this, they can modify the layer name
        If (blockIndex = curLayerHover) Then
        
            With m_NameRect
                .Left = xTextOffset - 2
                .Top = yTextOffset - 2
                .Right = xTextOffset + xTextWidth + 2
                .Bottom = yTextOffset + yTextHeight + 2
            End With
            
        End If
        
        'A few objects still need to be rendered below the current layer.  They all have the same y-offset, so calculate it in advance.
        yObjOffset = yTextOffset + layerNameFont.getHeightOfString(drawString) + 6
        
        'If this layer is currently hovered, draw some extra controls beneath the layer name.  This keeps the
        ' layer box from getting too cluttered, because we only draw relevant controls for the hovered layer.
        ' (Note that this approach is not touch-friendly; I'm aware, and will revisit as necessary if users
        '  request a touch-centric UI.)
        If (blockIndex = curLayerHover) Then
        
            'Start with an x-offset at the far right of the panel
            xObjOffset = m_BufferWidth - img_EyeClosed.getDIBWidth - fixDPI(DIST_BETWEEN_HOVER_BUTTONS)
        
            'Draw the visibility toggle.  Note that an icon for the opposite visibility state is drawn, to show
            ' the user what will happen if they click the icon.
            If tmpLayerRef.getLayerVisibility Then
                img_EyeClosed.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
            Else
                img_EyeOpen.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
            End If
            
            'Store the visibility toggle's rect (so that mouse events can more easily calculate hit events)
            fillRectWithDIBCoords m_VisibilityRect, img_EyeOpen, xObjOffset, yObjOffset
            
            'Next, provide a "duplicate layer" shortcut
            xObjOffset = xObjOffset - img_EyeOpen.getDIBWidth - fixDPI(DIST_BETWEEN_HOVER_BUTTONS)
            img_Duplicate.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
            fillRectWithDIBCoords m_DuplicateRect, img_Duplicate, xObjOffset, yObjOffset
            
            'Next, give the user dedicated merge down/up buttons.  These are only available if the layer is visible.
            If tmpLayerRef.getLayerVisibility Then
            
                'Merge down comes first...
                xObjOffset = xObjOffset - img_Duplicate.getDIBWidth - fixDPI(DIST_BETWEEN_HOVER_BUTTONS)
                
                If Layer_Handler.isLayerAllowedToMergeAdjacent(blockIndex, True) >= 0 Then
                    img_MergeDown.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
                Else
                    img_MergeDownDisabled.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
                End If
                fillRectWithDIBCoords m_MergeDownRect, img_MergeDown, xObjOffset, yObjOffset
                
                '...then Merge up
                xObjOffset = xObjOffset - img_MergeDown.getDIBWidth - fixDPI(DIST_BETWEEN_HOVER_BUTTONS)
                If Layer_Handler.isLayerAllowedToMergeAdjacent(blockIndex, False) >= 0 Then
                    img_MergeUp.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
                Else
                    img_MergeUpDisabled.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
                End If
                fillRectWithDIBCoords m_MergeUpRect, img_MergeUp, xObjOffset, yObjOffset
                
            End If
            
        End If
        
    End If

End Sub

'Given a destination rect and a UI DIB, fill the rect with the UI DIB's coordinates
Private Sub fillRectWithDIBCoords(ByRef dstRect As RECT, ByRef srcDIB As pdDIB, ByVal xOffset As Long, ByVal yOffset As Long)
    With dstRect
        .Left = xOffset
        .Top = yOffset
        .Right = xOffset + srcDIB.getDIBWidth
        .Bottom = yOffset + srcDIB.getDIBHeight
    End With
End Sub

'Given mouse coordinates over the buffer picture box, return the layer at that location.
' The optional parameter "reportNearestLayer" will return the index of the top layer if the mouse is in the invalid area
' above the top-most layer, and the bottom layer if in the invalid area beneath the bottom-most layer.
Private Function getLayerAtPosition(ByVal x As Long, ByVal y As Long, Optional ByVal reportNearestLayer As Boolean = False) As Long
    
    If pdImages(g_CurrentImage) Is Nothing Then
        getLayerAtPosition = -1
        Exit Function
    End If
    
    Dim vOffset As Long
    vOffset = vsLayer.Value
    
    Dim tmpLayerCheck As Long
    tmpLayerCheck = (y + vOffset) \ fixDPI(BLOCKHEIGHT)
    
    'It's a bit counterintuitive, but we draw the layer box in reverse order: layer 0 is at the BOTTOM,
    ' and layer(max) is at the TOP.  Because of this, all layer positioning checks must be reversed.
    tmpLayerCheck = (pdImages(g_CurrentImage).getNumOfLayers - 1) - tmpLayerCheck
    
    'Is the mouse over an actual layer, or just dead space in the box?
    If Not pdImages(g_CurrentImage) Is Nothing Then
    
        If (tmpLayerCheck >= 0) And (tmpLayerCheck < pdImages(g_CurrentImage).getNumOfLayers) Then
            getLayerAtPosition = tmpLayerCheck
        Else
        
            'If the user wants us to report the *nearest* valid layer
            If reportNearestLayer Then
            
                If tmpLayerCheck < 0 Then
                    getLayerAtPosition = 0
                Else
                    getLayerAtPosition = pdImages(g_CurrentImage).getNumOfLayers - 1
                End If
            
            'The user doesn't want us to report the nearest layer.  Report that the mouse is not over a layer.
            Else
                getLayerAtPosition = -1
            End If
            
        End If
    
    End If
    
End Function

Private Sub picLayers_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub
    
    'Use the external function (in the clipboard handler, as the code is roughly identical to clipboard pasting)
    ' to load the OLE source.
    m_InOLEDragDropMode = True
    Clipboard_Handler.loadImageFromDragDrop Data, Effect, True
    m_InOLEDragDropMode = False

End Sub

Private Sub picLayers_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub

    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Or Data.GetFormat(vbCFText) Then
        'Inform the source that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files or text, don't allow a drop
        Effect = vbDropEffectNone
    End If

End Sub

'Change the opacity of the current layer
Private Sub sltLayerOpacity_Change()

    'By default, changing the scroll bar will automatically update the opacity value of the selected layer, and
    ' the main viewport will be redrawn.  When changing the scrollbar programmatically, set disableRedraws to TRUE
    ' to prevent cylical redraws.
    If disableRedraws Then Exit Sub

    If g_OpenImageCount > 0 Then
    
        If Not pdImages(g_CurrentImage).getActiveLayer Is Nothing Then
        
            pdImages(g_CurrentImage).getActiveLayer.setLayerOpacity sltLayerOpacity.Value
            ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        End If
    
    End If

End Sub

Private Sub txtLayerName_KeyPress(KeyAscii As Integer)

    'KeyAscii 13 = Enter key; when this happens, commit the changed layer name and hide the text box
    If KeyAscii = 13 Then
        
        'Prevent beeps
        KeyAscii = 0
                
        'Set the active layer name, then hide the text box
        pdImages(g_CurrentImage).getActiveLayer.setLayerName txtLayerName.Text
        txtLayerName.Text = ""
        txtLayerName.Visible = False
        
        'Re-enable hotkeys now that editing is finished
        FormMain.ctlAccelerator.Enabled = True
        
        'Redraw the layer box with the new name
        redrawLayerBox
        
        'Transfer focus back to the layer box
        picLayers.SetFocus
        
    End If
    
End Sub

'If the text box loses focus mid-edit, hide it and discard any changes
Private Sub txtLayerName_LostFocus()

    'Hide the text box if it's still visible (e.g. if the user decided not to change a layer name after all).
    If txtLayerName.Visible Then txtLayerName.Visible = False

End Sub

Private Sub vsLayer_Change()
    redrawLayerBox
End Sub

Private Sub vsLayer_Scroll()
    redrawLayerBox
End Sub

'Whenever the layer toolbox is resized, we must reflow all objects to fill the available space.  Note that we do not do
' specialized handling for the vertical direction; vertically, the only change we handle is resizing the layer box itself
' to fill whatever vertical space is available.
Private Sub reflowInterface()

    'When the parent form is resized, resize the layer list (and other items) to properly fill the
    ' available horizontal and vertical space.
    
    'This value will be used to check for minimizing.  If the window is going down, we do not want to attempt a resize!
    Dim sizeCheck As Long
    
    'Start by moving the button box to the bottom of the available area
    sizeCheck = Me.ScaleHeight - picLayerButtons.Height - fixDPI(8)
    If sizeCheck > 0 Then picLayerButtons.Top = sizeCheck Else Exit Sub
    
    'Next, stretch the layer box to fill the available space
    sizeCheck = (picLayerButtons.Top - picLayers.Top) - fixDPI(8)
    If sizeCheck > 0 Then picLayers.Height = (picLayerButtons.Top - picLayers.Top) - fixDPI(8) Else Exit Sub
    
    'Make the toolbar the same height as the layer box
    vsLayer.Height = picLayers.Height
    
    'Vertical resizing has now been covered successfully.  Time to handle horizontal resizing.
    
    'Horizontally stretch the opacity and blend mode UI objects
    sltLayerOpacity.Width = Me.ScaleWidth - (sltLayerOpacity.Left + fixDPI(2))
    cboBlendMode.Width = Me.ScaleWidth - (cboBlendMode.Left + fixDPI(8))
    lnSeparator(0).x2 = Me.ScaleWidth - lnSeparator(0).x1
    
    'Resize the layer box and associated scrollbar
    vsLayer.Left = Me.ScaleWidth - vsLayer.Width - fixDPI(8)
    updateLayerScrollbarVisibility
       
    'Reflow the bottom button box; this is inevitably more complicated, owing to the spacing requirements of the buttons
    picLayerButtons.Width = Me.ScaleWidth
    
    '48px (at 96 DPI) is the ideal distance between buttons: 36px for the button, plus 12px for spacing.
    ' The total size of the button area of the box is thus 4 * 36 + 3 * 12, for FOUR buttons and THREE spacers.
    Dim buttonAreaWidth As Long, buttonAreaLeft As Long
    buttonAreaWidth = fixDPI(4 * 36 + 3 * 12)
    buttonAreaLeft = (Me.ScaleWidth - buttonAreaWidth) \ 2
    
    Dim i As Long
    For i = 0 To cmdLayerAction.Count - 1
        cmdLayerAction(i).Left = buttonAreaLeft + (i * 48)
    Next i
    
    'Redraw the internal layer UI DIB
    resizeLayerUI

End Sub
