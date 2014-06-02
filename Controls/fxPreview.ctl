VERSION 5.00
Begin VB.UserControl fxPreviewCtl 
   AccessKeys      =   "T"
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
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
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ToolboxBitmap   =   "fxPreview.ctx":0000
   Begin PhotoDemon.jcbutton cmdFit 
      Height          =   450
      Left            =   5160
      TabIndex        =   2
      Top             =   5160
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
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
      Mode            =   1
      Value           =   -1  'True
      HandPointer     =   -1  'True
      PictureNormal   =   "fxPreview.ctx":0312
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5100
      Left            =   0
      ScaleHeight     =   338
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   382
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5760
      Begin VB.VScrollBar vsOffsetY 
         Height          =   1335
         Left            =   5280
         TabIndex        =   4
         Top             =   3360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar hsOffsetX 
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   4680
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Label lblBeforeToggle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "show original image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07031&
      Height          =   210
      Left            =   120
      MouseIcon       =   "fxPreview.ctx":1064
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5280
      Width           =   1590
   End
End
Attribute VB_Name = "fxPreviewCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Effect Preview custom control
'Copyright ©2013-2014 by Tanner Helland
'Created: 10/January/13
'Last updated: 31/May/14
'Last update: convert custom mouse handling code to use pdInput
'
'For the first decade of its life, PhotoDemon relied on simple picture boxes for rendering its effect previews.
' This worked well enough when there were only a handful of tools available, but as the complexity of the program
' - and its various effects and tools - has grown, it has become more and more painful to update the preview
' system, because any changes have to be mirrored across a huge number of forms.
'
'Thus, this control was born.  It is now used on every single effect form in place of a regular picture box.  This
' allows me to add preview-related features just once - to the base control - and have every tool automatically
' reap the benefits.
'
'The control is capable of storing a copy of the original image and any filter-modified versions of the image.
' The user can toggle between these by using the command link below the main picture box, or by pressing Alt+T.
' This replaces the side-by-side "before and after" of past versions.
'
'A few other extra features have been implemented, which can be enabled on a tool-by-tool basis.  Specifically:
' 1) The user can toggle between "fit image" and "100% zoom + click-drag-to-scroll" modes.  Note that 100% zoom
'    is not appropriate for some tools (i.e. perspective transformations and other algorithms that only operate
'    on the full image area).
' 2) Click-to-select color functionality.  This is helpful for tools that rely on color information within the
'    image for their operation, e.g. green screen.
' 3) Click-to-select-coordinate functionality.  This is helpful for giving the user an easy way to select a
'    location on the image as, say, a center point for a filter (e.g. vignetting works great with this).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Preview boxes can now let the user switch between "full image" and "100% zoom" states
Public Event ViewportChanged()

'Some preview boxes will let the user click to set a new centerpoint for a filter or effect.
Public Event PointSelected(xRatio As Double, yRatio As Double)
Private isPointSelectionAllowed As Boolean

'Some preview boxes allow the user to click and select a color from the source image
Public Event ColorSelected()
Private isColorSelectionAllowed As Boolean, curColor As Long
Private colorJustClicked As Long

'Because some tools believe they are always operating on a full image (e.g. perspective transform), it may be necessary
' to disable zoom toggle on those controls
Private disableZoomPanAbility As Boolean

'Has this control been given a copy of the original image?
Private m_HasOriginal As Boolean, m_HasFX As Boolean

Private originalImage As pdDIB, fxImage As pdDIB

'The control's current state: whether it is showing the original image or the fx preview
Private curImageState As Boolean

'GetPixel is used to retrieve colors from the image
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

'Mouse events are raised with the help of the pdInput class
Private WithEvents cMouseEvents As pdInput
Attribute cMouseEvents.VB_VarHelpID = -1

'If the viewport is not set to "fit 100%", the user can click-drag around the image.  To do this successfully,
' we must track mouse position and offsets.
Private m_InitX As Long, m_InitY As Long
Private m_OffsetX As Long, m_OffsetY As Long
Private m_PrevOffsetX As Long, m_PrevOffsetY As Long

'Is the image large enough that the user is allowed to scroll?
Private m_HScrollAllowed As Boolean, m_VScrollAllowed As Boolean

Private Sub cmdFit_Click()
    
    'Note that we no longer have a valid copy of the original image data, so prepImageData must supply us with a new one
    m_HasOriginal = False
    m_HasFX = False
    
    'Raise a viewport change event so the containing form can redraw itself accordingly
    RaiseEvent ViewportChanged
    
End Sub

'If we don't expose an hWnd, any embedded jcButton controls will throw errors
Public Property Get offsetX() As Long
    If m_HScrollAllowed Then
        offsetX = validateXOffset(hsOffsetX.Value + m_OffsetX)
    Else
        offsetX = 0
    End If
End Property

'If we don't expose an hWnd, any embedded jcButton controls will throw errors
Public Property Get offsetY() As Long
    If m_VScrollAllowed Then
        offsetY = validateYOffset(vsOffsetY.Value + m_OffsetY)
    Else
        offsetY = 0
    End If
End Property

'If we don't expose an hWnd, any embedded jcButton controls will throw errors
Public Property Get viewportFitFullImage() As Boolean
    viewportFitFullImage = CBool(cmdFit.Value)
End Property

'If we don't expose an hWnd, any embedded jcButton controls will throw errors
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'External functions may need to access the color selected by the preview control
Public Property Get SelectedColor() As Long
    SelectedColor = curColor
End Property

'At design-time, use this property to determine whether the user is allowed to select colors directly from the
' preview window (helpful for tools like green screen, etc).
Public Property Get AllowColorSelection() As Boolean
    AllowColorSelection = isColorSelectionAllowed
End Property

Public Property Let AllowColorSelection(ByVal isAllowed As Boolean)
    isColorSelectionAllowed = isAllowed
    PropertyChanged "AllowColorSelection"
End Property

'At design-time, use this property to determine whether the user is allowed to select new center points for a filter
' or effect by clicking the preview window.
Public Property Get AllowPointSelection() As Boolean
    AllowPointSelection = isPointSelectionAllowed
End Property

Public Property Let AllowPointSelection(ByVal isAllowed As Boolean)
    isPointSelectionAllowed = isAllowed
    PropertyChanged "AllowPointSelection"
End Property

'At design-time, use this property to prevent the user from changing the preview area between zoom/pan and fit mode.
Public Property Get AllowZoomPan() As Boolean
    AllowZoomPan = Not disableZoomPanAbility
End Property

Public Property Let AllowZoomPan(ByVal isAllowed As Boolean)
    disableZoomPanAbility = Not isAllowed
    PropertyChanged "DisableZoomPan"
    redrawControl
    UserControl.Refresh
End Property

'Use this to supply the preview with a copy of the original image's data.  The preview object can use this to display
' the original image when the user clicks the "show original image" link.
Public Sub setOriginalImage(ByRef srcDIB As pdDIB)

    'Note that we have a copy of the original image, so the calling function doesn't attempt to supply it again
    m_HasOriginal = True
    
    'Make a copy of the DIB passed in
    If (originalImage Is Nothing) Then Set originalImage = New pdDIB
    
    originalImage.eraseDIB
    originalImage.createFromExistingDIB srcDIB
    
    If originalImage.getDIBColorDepth = 32 Then originalImage.fixPremultipliedAlpha True
    
End Sub

'Use this to supply the object with a copy of the processed image's data.  The preview object can use this to display
' the processed image again if the user clicks the "show original image" link, then clicks it again.
Public Sub setFXImage(ByRef srcDIB As pdDIB)

    'Note that we have a copy of the original image, so the calling function doesn't attempt to supply it again
    m_HasFX = True
    
    'Make a copy of the DIB passed in
    If (fxImage Is Nothing) Then Set fxImage = New pdDIB
    
    fxImage.eraseDIB
    fxImage.createFromExistingDIB srcDIB
        
    'If the user was previously examining the original image, and color selection is not allowed, be helpful and
    ' automatically restore the previewed image.
    If (Not isColorSelectionAllowed) Then
        fxImage.renderToPictureBox picPreview
        lblBeforeToggle.Caption = g_Language.TranslateMessage("show original image") & " (alt+t) "
        curImageState = True
    'If color selection is allowed, the user may want to select more colors - so leave it on "original" mode if it
    ' is already there.
    Else
        If curImageState Then
            fxImage.renderToPictureBox picPreview
            lblBeforeToggle.Caption = g_Language.TranslateMessage("show original image") & " (alt+t) "
        End If
    End If

End Sub

'Has this preview control had an original version of the image set?
Public Function hasOriginalImage() As Boolean
    hasOriginalImage = m_HasOriginal
End Function

'Return a handle to our primary picture box
Public Function getPreviewPic() As PictureBox
    Set getPreviewPic = picPreview
End Function

'Return dimensions of the preview picture box
Public Function getPreviewWidth() As Long
    getPreviewWidth = picPreview.ScaleWidth
End Function

Public Function getPreviewHeight() As Long
    getPreviewHeight = picPreview.ScaleHeight
End Function

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'If this preview control instance allows the user to select a color, display the original image upon mouse entrance
    If viewportFitFullImage Then
        If AllowColorSelection Then
            cMouseEvents.setPNGCursor "C_PIPETTE", 0, 0
            If (Not originalImage Is Nothing) Then originalImage.renderToPictureBox picPreview
        End If
    Else
        cMouseEvents.setSystemCursor IDC_HAND
    End If

End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'If this preview control instance allows the user to select a color, restore whatever image was previously
    ' displayed upon mouse exit
    If AllowColorSelection Then
        
        cMouseEvents.setSystemCursor IDC_HAND
        
        If curImageState Then
            If (Not fxImage Is Nothing) Then fxImage.renderToPictureBox picPreview
        Else
            If (Not originalImage Is Nothing) Then originalImage.renderToPictureBox picPreview
        End If
    End If

End Sub

'Toggle between the preview image and the original image if the user clicks this label
Private Sub lblBeforeToggle_Click()
    
    'Before doing anything else, change the label caption
    If curImageState Then
        lblBeforeToggle.Caption = g_Language.TranslateMessage("show effect preview") & " (alt+t) "
    Else
        lblBeforeToggle.Caption = g_Language.TranslateMessage("show original image") & " (alt+t) "
    End If
    lblBeforeToggle.Refresh
    
    curImageState = Not curImageState
    
    'Update the image to match the new caption
    If Not curImageState Then
        If m_HasOriginal Then originalImage.renderToPictureBox picPreview
    Else
        
        If m_HasFX Then
            fxImage.renderToPictureBox picPreview
        Else
            If m_HasOriginal Then originalImage.renderToPictureBox picPreview
        End If
    End If
    
End Sub

'If color selection is allowed, raise that event now
Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If viewport scrolling is allowed, initialize it now
    If Not viewportFitFullImage Then
        If Button = vbLeftButton Then
            m_InitX = x
            m_InitY = y
            cMouseEvents.setSystemCursor IDC_SIZEALL
        End If
    End If
    
    'If color selection is allowed, initialize it now
    If isColorSelectionAllowed Then
        
        If Button = vbRightButton Then
        
            curColor = GetPixel(originalImage.getDIBDC, x - ((picPreview.ScaleWidth - originalImage.getDIBWidth) \ 2), y - ((picPreview.ScaleHeight - originalImage.getDIBHeight) \ 2))
            
            If curColor = -1 Then curColor = RGB(127, 127, 127)
            
            If AllowColorSelection Then colorJustClicked = 1
            RaiseEvent ColorSelected
            
        End If
        
    End If
    
    'If point selection is allowed, initialize it now
    If isPointSelectionAllowed Then
    
        If (Button = vbRightButton) Or (Button = vbLeftButton) Then
        
            'Return the mouse coordinates as a ratio between 0 and 1, with 1 representing max width/height
            Dim retX As Double, retY As Double
            retX = x - ((picPreview.ScaleWidth - originalImage.getDIBWidth) \ 2)
            retY = y - ((picPreview.ScaleHeight - originalImage.getDIBHeight) \ 2)
            
            retX = retX / originalImage.getDIBWidth
            retY = retY / originalImage.getDIBHeight
            
            RaiseEvent PointSelected(retX, retY)
        
        End If
    
    End If
    
End Sub

'When the user is selecting a color, we want to give them a preview of how that color will affect the previewed image.
' This is handled in the _MouseDown event above.  After the color has been selected, we want to restore the original
' image on a subsequent mouse move, in case the user wants to select a different color.
Private Sub picPreview_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If the viewport is not set to "fit to screen", then we must determine offsets based on the mouse position
    If Not viewportFitFullImage Then
        
        If Button = vbLeftButton Then
        
            'Make sure the move cursor remains accurate
            cMouseEvents.setSystemCursor IDC_SIZEALL
                
            'Store new offsets for the image
            m_OffsetX = m_InitX - x
            m_OffsetY = m_InitY - y
            
            'Note that we no longer have a valid copy of the original image data, so prepImageData must supply us with a new one
            m_HasOriginal = False
            m_HasFX = False
            
            'Raise an external viewport change event that tool dialogs can use to refresh their effect preview
            RaiseEvent ViewportChanged
            
        Else
            If Not isColorSelectionAllowed Then cMouseEvents.setSystemCursor IDC_HAND
        End If
    Else
        'setArrowCursor picPreview
    End If
    
    If colorJustClicked > 0 Then
    
        'To accomodate shaky hands, allow a few mouse movements before resetting the image
        If colorJustClicked < 4 Then
            colorJustClicked = colorJustClicked + 1
        Else
            colorJustClicked = 0
            If (Not originalImage Is Nothing) Then originalImage.renderToPictureBox picPreview
        End If
        
    End If
    
    'If point selection is allowed, continue firing events while the mouse is moving (as a convenience to the user)
    If isPointSelectionAllowed Then
    
        cMouseEvents.setSystemCursor IDC_HAND
    
        If (Button = vbRightButton) Or (Button = vbLeftButton) Then
        
            'Return the mouse coordinates as a ratio between 0 and 1, with 1 representing max width/height
            Dim retX As Double, retY As Double
            retX = x - ((picPreview.ScaleWidth - originalImage.getDIBWidth) \ 2)
            retY = y - ((picPreview.ScaleHeight - originalImage.getDIBHeight) \ 2)
            
            retX = retX / originalImage.getDIBWidth
            retY = retY / originalImage.getDIBHeight
            
            RaiseEvent PointSelected(retX, retY)
        
        End If
    
    End If
    
End Sub

Private Sub picPreview_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Not viewportFitFullImage Then
        
        cMouseEvents.setSystemCursor IDC_HAND
        
        hsOffsetX.Value = validateXOffset(hsOffsetX.Value + m_OffsetX)
        m_OffsetX = 0
        
        vsOffsetY.Value = validateYOffset(vsOffsetY.Value + m_OffsetY)
        m_OffsetY = 0
        
    End If

End Sub

'X and Y offsets for the image preview are generated dynamically by the user's mouse movements.  As multiple functions
' need to validate those offsets to make sure they don't result in an offset outside the image, these standardized
' validation functions were created.
Private Function validateXOffset(ByVal currentOffset As Long) As Long
    If currentOffset < 0 Then currentOffset = 0
    If currentOffset > hsOffsetX.Max Then currentOffset = hsOffsetX.Max
    validateXOffset = currentOffset
End Function

Private Function validateYOffset(ByVal currentOffset As Long) As Long
    If currentOffset < 0 Then currentOffset = 0
    If currentOffset > vsOffsetY.Max Then currentOffset = vsOffsetY.Max
    validateYOffset = currentOffset
End Function

'I haven't made up my mind on whether to use AutoRedraw or not; just to be safe, I've added handling code to the _Paint
' event so that AutoRedraw can be turned off without trouble.
Private Sub picPreview_Paint()

    'Update the image to match the before/after label state
    If Not curImageState Then
        If m_HasOriginal Then originalImage.renderToPictureBox picPreview
    Else
        
        If m_HasFX Then
            fxImage.renderToPictureBox picPreview
        Else
            If m_HasOriginal Then originalImage.renderToPictureBox picPreview
        End If
    End If

End Sub

'When the control's access key is pressed (alt+t) , toggle the original/current image
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    lblBeforeToggle_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    
    'Keep the control's backcolor in sync with the parent object
    If UCase$(PropertyName) = "BACKCOLOR" Then
        BackColor = Ambient.BackColor
    End If

End Sub

Private Sub UserControl_Initialize()
    
    'A check must be made for IDE behavior so the project will compile; VB's initialization of user controls during
    ' compiling and design process causes no shortage of odd issues and errors otherwise
    If g_UserModeFix Then
        
        'Set up a mouse events handler.  (NOTE: this handler subclasses, which may cause instability in the IDE.)
        Set cMouseEvents = New pdInput
        cMouseEvents.addInputTracker picPreview.hWnd, True, , , True
        cMouseEvents.setSystemCursor IDC_ARROW
        
        'Give the toggle image text the same font as the rest of the project.
        lblBeforeToggle.FontName = g_InterfaceFont
        
    End If
    
    curImageState = True
    curColor = 0
            
End Sub

'Initialize our effect preview control
Private Sub UserControl_InitProperties()
    
    'Set the background of the fxPreview to match the background of our parent object
    BackColor = Ambient.BackColor
    
    'Mark the original image as having NOT been set
    m_HasOriginal = False
    
    'By default, the control cannot be used for color selection
    isColorSelectionAllowed = False
    
    'By default, the control allows the user to zoom/pan the transformation
    disableZoomPanAbility = False
    
    'By default, the control does not allow for selecting coordinate points by clicking
    isPointSelectionAllowed = False
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        AllowColorSelection = .ReadProperty("ColorSelection", False)
        AllowPointSelection = .ReadProperty("PointSelection", False)
        disableZoomPanAbility = .ReadProperty("DisableZoomPan", False)
    End With
    
End Sub

'Redraw the user control after it has been resized
Private Sub UserControl_Resize()
    redrawControl
End Sub

Private Sub UserControl_Show()
    
    'Translate the user control text in the compiled EXE
    If g_UserModeFix Then
        lblBeforeToggle.Caption = g_Language.TranslateMessage("show original image") & " (alt+t) "
    Else
        lblBeforeToggle.Caption = "show original image (alt+t) "
    End If
        
    'setArrowCursorToHwnd UserControl.hWnd
    
    'Ensure the control is redrawn at least once
    redrawControl
    
    'Set an initial max/min for the preview offsets if the user chooses to preview at 100% zoom
    If g_UserModeFix Then
    
        'Reset the mouse cursor
        cMouseEvents.setSystemCursor IDC_ARROW
    
        Dim maxHOffset As Long, maxVOffset As Long
        
        Dim srcWidth As Long, srcHeight As Long
        If pdImages(g_CurrentImage).selectionActive Then
            srcWidth = pdImages(g_CurrentImage).mainSelection.boundWidth
            srcHeight = pdImages(g_CurrentImage).mainSelection.boundHeight
        Else
            srcWidth = pdImages(g_CurrentImage).getActiveDIB.getDIBWidth
            srcHeight = pdImages(g_CurrentImage).getActiveDIB.getDIBHeight
        End If
        
        maxHOffset = srcWidth - picPreview.ScaleWidth
        maxVOffset = srcHeight - picPreview.ScaleHeight
        
        If maxHOffset > 0 Then
            hsOffsetX.Max = maxHOffset
            m_HScrollAllowed = True
        Else
            hsOffsetX.Max = 1
            m_HScrollAllowed = False
        End If
        
        If maxVOffset > 0 Then
            vsOffsetY.Max = maxVOffset
            m_VScrollAllowed = True
        Else
            vsOffsetY.Max = 1
            m_VScrollAllowed = False
        End If
    
    End If
    
End Sub

Private Sub UserControl_Terminate()

    'Release any image objects that may have been created
    If Not (originalImage Is Nothing) Then originalImage.eraseDIB
    If Not (fxImage Is Nothing) Then fxImage.eraseDIB
    
End Sub

'After a resize or paint request, update the layout of our control
Private Sub redrawControl()
    
    'Always make the preview picture box the width of the user control (at present)
    picPreview.Width = UserControl.ScaleWidth
    
    'Adjust the preview picture box's height to be just above the "show original image" link
    lblBeforeToggle.Top = UserControl.ScaleHeight - fixDPI(24)
    picPreview.Height = lblBeforeToggle.Top - (UserControl.ScaleHeight - (lblBeforeToggle.Height + lblBeforeToggle.Top))
    
    'Align the fit/100% toggle button
    'cmdFit.Left = UserControl.ScaleWidth - cmdFit.Width
    'cmdFit.Top = picPreview.Height + ((UserControl.ScaleHeight - (picPreview.Height + cmdFit.Height)) / 2)
    cmdFit.Height = UserControl.ScaleHeight - picPreview.Height - (fixDPI(2) * 2)
    cmdFit.Width = cmdFit.Height
    cmdFit.Top = picPreview.Height + fixDPI(2)
    cmdFit.Left = UserControl.ScaleWidth - cmdFit.Width '- fixDPI(2)
    cmdFit.forceButtonRedraw
    
    'If zoom/pan is not allowed, hide that button entirely
    If disableZoomPanAbility Then cmdFit.Visible = False Else cmdFit.Visible = True
        
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "ColorSelection", AllowColorSelection, False
        .WriteProperty "DisableZoomPan", disableZoomPanAbility, False
        .WriteProperty "PointSelection", AllowPointSelection, False
    End With
    
End Sub
