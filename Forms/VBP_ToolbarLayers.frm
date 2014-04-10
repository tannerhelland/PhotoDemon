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
   Begin VB.VScrollBar vsLayer 
      Height          =   4905
      LargeChange     =   32
      Left            =   3360
      Max             =   100
      TabIndex        =   9
      Top             =   1200
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
      TabIndex        =   5
      Top             =   6750
      Width           =   3735
      Begin PhotoDemon.jcbutton cmdLayerAction 
         Height          =   480
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   15
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   847
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
         PictureNormal   =   "VBP_ToolbarLayers.frx":0000
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
      End
      Begin PhotoDemon.jcbutton cmdLayerAction 
         Height          =   480
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   15
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   847
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
         PictureNormal   =   "VBP_ToolbarLayers.frx":0D52
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
      End
      Begin PhotoDemon.jcbutton cmdLayerAction 
         Height          =   480
         Index           =   2
         Left            =   2160
         TabIndex        =   8
         Top             =   15
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   847
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
         PictureNormal   =   "VBP_ToolbarLayers.frx":1AA4
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
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
      ScaleHeight     =   327
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   215
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "Layer Name"
      Top             =   105
      Width           =   2535
   End
   Begin PhotoDemon.sliderTextCombo sltLayerOpacity 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   873
      Max             =   100
      Value           =   100
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
   Begin VB.Line lnSeparator 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   8
      X2              =   240
      Y1              =   72
      Y2              =   72
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
      TabIndex        =   2
      Top             =   570
      Width           =   675
   End
   Begin VB.Label lblLayerSettings 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "name:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "toolbar_Layers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private WithEvents cMouseEvents As bluMouseEvents
Attribute cMouseEvents.VB_VarHelpID = -1

'Height of each layer content block.  Note that this is effectively a "magic number", in pixels, representing the
' height of each layer block in the layer selection UI.  This number will be dynamically resized per the current
' screen DPI by the "redrawLayerList" and "renderLayerBlock" functions.
Private Const BLOCKHEIGHT As Long = 48

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
    LYR_BTN_MOVE_UP = 0
    LYR_BTN_MOVE_DOWN = 1
    LYR_BTN_DELETE = 2
End Enum

#If False Then
    Private Const LYR_BTN_MOVE_UP = 0, LYR_BTN_MOVE_DOWN = 1, LYR_BTN_DELETE = 2
#End If

'Sometimes we need to make changes that will raise redraw-causing events.  Set this variable to TRUE if you want
' such functions to ignore their automatic redrawing.
Private disableRedraws As Boolean

'Extra interface images are loaded as resources at run-time
Private img_EyeOpen As pdDIB, img_EyeClosed As pdDIB

'Some UI elements are dynamically rendered onto the layer box.  To simplify hit detection, their RECTs are stored
' at render-time, which allows the mouse actions to easily check hits regardless of layer box position.
Private m_VisibilityRect As RECT

'External functions can force a full redraw by calling this sub.  (This is necessary whenever layers are added, deleted,
' re-ordered, etc.)
Public Sub forceRedraw(Optional ByVal refreshThumbnailCache As Boolean = True)
    
    If refreshThumbnailCache Then cacheLayerThumbnails
    
    'Synchronize the opacity scroll bar to the active layer.
    disableRedraws = True
    If (g_OpenImageCount > 0) Then
        If (Not pdImages(g_CurrentImage).getActiveLayer Is Nothing) Then
            sltLayerOpacity.Value = pdImages(g_CurrentImage).getActiveLayer.getLayerOpacity
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

'Layer action buttons - move layers up/down, delete layers, etc.
Private Sub cmdLayerAction_Click(Index As Integer)

    Dim copyOfCurLayerID As Long
    copyOfCurLayerID = pdImages(g_CurrentImage).getActiveLayerID

    Select Case Index
    
        Case LYR_BTN_MOVE_UP
            pdImages(g_CurrentImage).moveLayerByIndex pdImages(g_CurrentImage).getActiveLayerIndex, True
            cacheLayerThumbnails
            Layer_Handler.setActiveLayerByID copyOfCurLayerID, True
        
        Case LYR_BTN_MOVE_DOWN
            pdImages(g_CurrentImage).moveLayerByIndex pdImages(g_CurrentImage).getActiveLayerIndex, False
            cacheLayerThumbnails
            Layer_Handler.setActiveLayerByID copyOfCurLayerID, True
    
        Case LYR_BTN_DELETE
            pdImages(g_CurrentImage).deleteLayerByIndex pdImages(g_CurrentImage).getActiveLayerIndex
            cacheLayerThumbnails
            Layer_Handler.setActiveLayerByIndex pdImages(g_CurrentImage).getActiveLayerIndex, True
    
    End Select
    
End Sub

Private Sub cMouseEvents_MouseOut()
    curLayerHover = -1
    redrawLayerBox
End Sub

Private Sub cMouseEvents_MouseVScroll(ByVal LinesScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)

    'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
    If vsLayer.Visible Then
  
        If LinesScrolled < 0 Then
            
            If vsLayer.Value + vsLayer.LargeChange > vsLayer.Max Then
                vsLayer.Value = vsLayer.Max
            Else
                vsLayer.Value = vsLayer.Value + vsLayer.LargeChange
            End If
            
            curLayerHover = getLayerAtPosition(x, y)
            redrawLayerBox
        
        ElseIf LinesScrolled > 0 Then
            
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

Private Sub Form_Load()

    'Reset the thumbnail array
    numOfThumbnails = 0
    ReDim layerThumbnails(0 To numOfThumbnails) As thumbEntry

    'Activate the custom tooltip handler
    Set m_ToolTip = New clsToolTip
    m_ToolTip.Create Me
    m_ToolTip.MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
    m_ToolTip.DelayTime(ttDelayShow) = 5000
    m_ToolTip.AddTool picLayers, ""
    
    'Theme the form
    makeFormPretty Me
    
    'Enable mousewheel scrolling for the layer box
    Set cMouseEvents = New bluMouseEvents
    cMouseEvents.Attach picLayers.hWnd, Me.hWnd
    cMouseEvents.MousePointer = IDC_HAND
    
    'No layer has been hovered yet
    curLayerHover = -1
    
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
    Set img_EyeOpen = New pdDIB
    loadResourceToDIB "EYE_OPEN", img_EyeOpen
    Set img_EyeClosed = New pdDIB
    loadResourceToDIB "EYE_CLOSE", img_EyeClosed
    
    'Pad all interface images with 2px blank space; this makes them a bit more aesthetically pleasing, and saves us the
    ' trouble of manually calculating 2px offsets for each image at draw-time
    Filters_Layers.padDIB img_EyeOpen, 2
    Filters_Layers.padDIB img_EyeClosed, 2
    
End Sub

Private Sub Form_Resize()

    'When the parent form is resized, resize the layer list (and other items) to properly fill the
    ' available vertical space.
    
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
    
    'Redraw the internal layer UI DIB
    resizeLayerUI

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
        toggleToolbarVisibility TOOLS_TOOLBOX
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
    thumbHeight = BLOCKHEIGHT - 2
    thumbWidth = thumbHeight
    
    'Redraw the toolbar
    redrawLayerBox
    
End Sub

'Cache all current layer thumbnails.  This is required for things like the user switching to a new image, which requires
' us to wipe the current layer cache and start anew.
Private Sub cacheLayerThumbnails()

    If Not pdImages(g_CurrentImage) Is Nothing Then
    
        'Make sure the active image has at least one layer
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
        
        End If
    
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
        If tmpLayerRef.getLayerVisibility Then
            layerThumbnails(blockIndex).thumbDIB.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
        Else
            layerThumbnails(blockIndex).thumbDIB.alphaBlendToDC bufferDIB.getDIBDC, 76, xObjOffset, yObjOffset
            
            'Also, render a "closed eye" icon in the corner.
            ' NOTE: I'm not sold on this being a good idea.  The icon seems to be clickable, but it isn't!
            'img_EyeClosed.alphaBlendToDC bufferDIB.getDIBDC, 210, xObjOffset + (BLOCKHEIGHT - img_EyeClosed.getDIBWidth) - fixDPI(5), yObjOffset + (BLOCKHEIGHT - img_EyeClosed.getDIBHeight) - fixDPI(6)
            
        End If
        
        'Render the layer name
        Dim drawString As String
        drawString = tmpLayerRef.getLayerName
        
        'If this layer is invisible, mark it as such.
        ' NOTE: not sold on this behavior, but I'm leaving it for a bit to see how it affects workflow.
        If Not tmpLayerRef.getLayerVisibility Then drawString = g_Language.TranslateMessage("(hidden)") & " " & drawString
        
        layerNameFont.attachToDC bufferDIB.getDIBDC
        
        Dim xTextOffset As Long, yTextOffset As Long
        xTextOffset = offsetX + thumbWidth + fixDPI(thumbBorder) * 2
        yTextOffset = offsetY + fixDPI(4)       'yTextOffset = offsetY + (BLOCKHEIGHT - layerNameFont.getHeightOfString(drawString) - fixDPI(2)) / 2  'This line of code can be used to center the text in the layer area.
        layerNameFont.fastRenderTextWithClipping xTextOffset, yTextOffset, m_BufferWidth - xTextOffset - fixDPI(4), layerNameFont.getHeightOfString(drawString), drawString
        
        'A few objects still need to be rendered below the current layer.  They all have the same y-offset, so calculate it in advance.
        yObjOffset = yTextOffset + layerNameFont.getHeightOfString(drawString) + fixDPI(4)
        
        'If this layer is currently hovered, draw some extra controls beneath the layer name.  This keeps the
        ' layer box from getting too cluttered, because we only draw relevant controls for the hovered layer.
        ' (Note that this approach is not touch-friendly; I'm aware, and will revisit as necessary if users
        '  request a touch-centric UI.)
        If (blockIndex = curLayerHover) Then
        
            'Start with an x-offset that matches the text offset
            xObjOffset = xTextOffset
        
            'Draw the visibility toggle.  Note that an icon for the opposite visibility state is drawn, to show
            ' the user what will happen if they click the icon.
            If tmpLayerRef.getLayerVisibility Then
                img_EyeClosed.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
            Else
                img_EyeOpen.alphaBlendToDC bufferDIB.getDIBDC, 255, xObjOffset, yObjOffset
            End If
            
            'Store the visibility toggle's rect (so that mouse events can more easily calculate hit events)
            With m_VisibilityRect
                .Left = xObjOffset
                .Top = yObjOffset
                .Right = xObjOffset + img_EyeOpen.getDIBWidth
                .Bottom = yObjOffset + img_EyeOpen.getDIBHeight
            End With
            
        End If
        
    End If

End Sub

'Layer box was clicked; set that layer as the new active layer, and notify the parent pdImage object
Private Sub picLayers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim clickedLayer As Long
    clickedLayer = getLayerAtPosition(x, y)
    
    If clickedLayer >= 0 Then
        
        If Not pdImages(g_CurrentImage) Is Nothing Then
            
            'Check the clicked position against a series of rects, each one representing a unique interaction.
            
            'Has the user clicked a visibility rectangle?
            If isPointInRect(x, y, m_VisibilityRect) Then
            
                Layer_Handler.setLayerVisibilityByIndex clickedLayer, Not pdImages(g_CurrentImage).getLayerByIndex(clickedLayer).getLayerVisibility, True
            
            'The user has not clicked any item of interest.  Assume that they want to make the clicked layer
            ' the active layer.
            Else
                Layer_Handler.setActiveLayerByIndex clickedLayer
            End If
            
            'Redraw the layer box to represent any changes from this interaction.
            ' NOTE: this is not currently necessary, as all interactions will automatically force a redraw on their own.
            'redrawLayerBox
            
        End If
        
    End If
    
End Sub

Private Sub picLayers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
            
    'Update the tooltip contingent on the mouse position.
    Dim toolString As String
    
    'Mouse is over a visibility toggle
    If isPointInRect(x, y, m_VisibilityRect) Then
            
        If pdImages(g_CurrentImage).getLayerByIndex(curLayerHover).getLayerVisibility Then
            toolString = g_Language.TranslateMessage("Click to hide this layer.")
        Else
            toolString = g_Language.TranslateMessage("Click to show this layer.")
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
    
    
    'If a layer other than the active one is being hovered, highlight that box
    If curLayerHover <> getLayerAtPosition(x, y) Then
        curLayerHover = getLayerAtPosition(x, y)
        redrawLayerBox
    End If
    
End Sub

'Given mouse coordinates over the buffer picture box, return the layer at that location
Private Function getLayerAtPosition(ByVal x As Long, ByVal y As Long) As Long
    
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
            getLayerAtPosition = -1
        End If
    
    End If
    
End Function

'Change the opacity of the current layer
Private Sub sltLayerOpacity_Change()

    'By default, changing the scroll bar will automatically update the opacity value of the selected layer, and
    ' the main viewport will be redrawn.  When changing the scrollbar programmatically, set disableRedraws to FALSE
    ' to prevent cylical redraws.
    If disableRedraws Then Exit Sub

    If g_OpenImageCount > 0 Then
    
        If Not pdImages(g_CurrentImage).getActiveLayer Is Nothing Then
        
            pdImages(g_CurrentImage).getActiveLayer.setLayerOpacity sltLayerOpacity.Value
            ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        End If
    
    End If

End Sub

Private Sub vsLayer_Change()
    redrawLayerBox
End Sub

Private Sub vsLayer_Scroll()
    redrawLayerBox
End Sub
