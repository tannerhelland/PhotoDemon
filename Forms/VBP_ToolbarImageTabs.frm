VERSION 5.00
Begin VB.Form toolbar_ImageTabs 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Images"
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13710
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
   NegotiateMenus  =   0   'False
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   914
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsThumbnails 
      Height          =   255
      Left            =   0
      Max             =   10
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   13695
   End
End
Attribute VB_Name = "toolbar_ImageTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Selection ("Tab") Toolbar
'Copyright ©2013-2014 by Tanner Helland
'Created: 15/October/13
'Last updated: 21/October/13
'Last update: fix syncing of curThumb and g_CurrentImage when an inactive image is unloaded via the taskbar
'
'In fall 2014, PhotoDemon left behind the MDI model in favor of fully dockable/floatable tool and image windows.
' This required quite a new features, including a way to switch between loaded images when image windows are docked -
' which is where this form comes in.
'
'The purpose of this form is to provide a tab-like interface for switching between open images.  Please note that
' much of this form's layout and alignment is handled by PhotoDemon's window manager, so you will need to look
' there for additional details.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'A collection of all currently active thumbnails; this is dynamically resized as thumbnails are added/removed.
Private Type thumbEntry
    thumbLayer As pdLayer
    thumbShadow As pdLayer
    indexInPDImages As Long
End Type

Private imgThumbnails() As thumbEntry
Private numOfThumbnails As Long

'Because the user can resize the thumbnail bar, we must track thumbnail width/height dynamically
Private thumbWidth As Long, thumbHeight As Long

'We don't want thumbnails to fill the full size of their individual blocks, so we apply a border of this many pixels
' to each side of the thumbnail
Private Const thumbBorder As Long = 5

'The back buffer we use to hold the thumbnail display
Private bufferLayer As pdLayer
Private m_BufferWidth As Long, m_BufferHeight As Long

'An outside class provides access to mousewheel events for scrolling the filter view
Private WithEvents cMouseEvents As bluMouseEvents
Attribute cMouseEvents.VB_VarHelpID = -1

'The currently selected and currently hovered thumbnail
Private curThumb As Long, curThumbHover As Long

'We allow the user to resize this window via the bottom border; these constants are used with the SendMessage API to enable this behavior
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTLEFT As Long = 10
Private Const HTTOP As Long = 12
Private Const HTRIGHT As Long = 11
Private Const HTBOTTOM As Long = 15

'When we are responsible for this window resizing (because the user is resizing our window manually), we set this to TRUE.
' This variable is then checked before requesting additional redraws during our resize event.
Private weAreResponsibleForResize As Boolean

'As a convenience to the user, we provide a small notification when an image has unsaved changes
Private unsavedChangesLayer As pdLayer

'Drop-shadows on the thumbnails have a variable radius that changes based on the user's DPI settings
Private shadowBlurRadius As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'If the user loads tons of images, the tabstrip may overflow the available area.  We now allow them to drag-scroll the list.
' In order to allow that, we must track a few extra things, like initial mouse x/y
Private m_MouseDown As Boolean, m_ScrollingOccured As Boolean
Private m_InitX As Long, m_InitY As Long, m_InitOffset As Long
Private m_ListScrollable As Boolean
Private m_MouseDistanceTraveled As Long

'Horizontal or vertical layout; obviously, all our rendering and mouse detection code changes depending on the orientation
' of the tabstrip.
Private verticalLayout As Boolean

'When resizing, it is almost certain that the user will move the mouse outside the form.  Track this, and use it to notify
' the mouse handler that the user is not click-dragging the image list.
Public nowResizing As Boolean

'External functions can force a full redraw by calling this sub
Public Sub forceRedraw()
    Form_Resize
End Sub

'When the user switches images, redraw the toolbar to match the change
Public Sub notifyNewActiveImage(ByVal newPDImageIndex As Long)
    
    'Find the matching thumbnail entry, and mark it as active
    Dim i As Long
    For i = 0 To numOfThumbnails
        If imgThumbnails(i).indexInPDImages = newPDImageIndex Then
            curThumb = i
            Exit For
        End If
    Next i
    
    'Redraw the toolbar to reflect the change
    redrawToolbar
        
End Sub

'When the user somehow changes an image, they need to notify the toolbar, so that a new thumbnail can be rendered
Public Sub notifyUpdatedImage(ByVal pdImagesIndex As Long)
    
    'Find the matching thumbnail entry, and update its thumbnail layer
    Dim i As Long
    For i = 0 To numOfThumbnails
        If imgThumbnails(i).indexInPDImages = pdImagesIndex Then
            
            If verticalLayout Then
                pdImages(pdImagesIndex).requestThumbnail imgThumbnails(i).thumbLayer, fixDPI(thumbHeight) - (fixDPI(thumbBorder) * 2)
            Else
                pdImages(pdImagesIndex).requestThumbnail imgThumbnails(i).thumbLayer, fixDPI(thumbWidth) - (fixDPI(thumbBorder) * 2)
            End If
            
            updateShadowLayer i
            Exit For
        End If
    Next i
    
    'Redraw the toolbar to reflect the change
    redrawToolbar
        
End Sub

'Whenever a new image is loaded, it needs to be registered with the toolbar
Public Sub registerNewImage(ByVal pdImagesIndex As Long)
    
    'Request a thumbnail from the relevant pdImage object, and premultiply it to allow us to blit it more quickly
    Set imgThumbnails(numOfThumbnails).thumbLayer = New pdLayer
    
    If verticalLayout Then
        pdImages(pdImagesIndex).requestThumbnail imgThumbnails(numOfThumbnails).thumbLayer, fixDPI(thumbHeight) - (fixDPI(thumbBorder) * 2)
    Else
        pdImages(pdImagesIndex).requestThumbnail imgThumbnails(numOfThumbnails).thumbLayer, fixDPI(thumbWidth) - (fixDPI(thumbBorder) * 2)
    End If
    
    'Create a matching shadow layer
    Set imgThumbnails(numOfThumbnails).thumbShadow = New pdLayer
    updateShadowLayer numOfThumbnails
    
    'Make a note of this thumbnail's index in the main pdImages array
    imgThumbnails(numOfThumbnails).indexInPDImages = pdImagesIndex
    
    'We can assume this image will be the active one
    curThumb = numOfThumbnails
    
    'Prepare the array to receive another entry in the future
    numOfThumbnails = numOfThumbnails + 1
    ReDim Preserve imgThumbnails(0 To numOfThumbnails) As thumbEntry
    
    'Redraw the toolbar to reflect these changes
    redrawToolbar
    
End Sub

'Whenever an image is unloaded, it needs to be de-registered with the toolbar
Public Sub RemoveImage(ByVal pdImagesIndex As Long)

    'Find the matching thumbnail in our collection
    Dim i As Long, thumbIndex As Long
    For i = 0 To numOfThumbnails
        If imgThumbnails(i).indexInPDImages = pdImagesIndex Then
            thumbIndex = i
            Exit For
        End If
    Next i
    
    'thumbIndex is now equal to the matching thumbnail.  Remove that entry, then shift all thumbnails after that point down.
    If Not (imgThumbnails(thumbIndex).thumbLayer Is Nothing) Then
        imgThumbnails(thumbIndex).thumbLayer.eraseLayer
        Set imgThumbnails(thumbIndex).thumbLayer = Nothing
        imgThumbnails(thumbIndex).thumbShadow.eraseLayer
        Set imgThumbnails(thumbIndex).thumbShadow = Nothing
    End If
    
    For i = thumbIndex To numOfThumbnails - 1
        Set imgThumbnails(i).thumbLayer = imgThumbnails(i + 1).thumbLayer
        Set imgThumbnails(i).thumbShadow = imgThumbnails(i + 1).thumbShadow
        imgThumbnails(i).indexInPDImages = imgThumbnails(i + 1).indexInPDImages
    Next i
    
    'Decrease the array size to erase the unneeded trailing entry
    numOfThumbnails = numOfThumbnails - 1
    If numOfThumbnails < 0 Then numOfThumbnails = 0
    ReDim Preserve imgThumbnails(0 To numOfThumbnails) As thumbEntry
    
    'Because inactive images can be unloaded via the Win 7 taskbar, it is possible for our curThumb tracker to get out of sync.
    ' Update it now.
    For i = 0 To numOfThumbnails
        If imgThumbnails(i).indexInPDImages = g_CurrentImage Then
            curThumb = i
            Exit For
        End If
    Next i
    
    'Redraw the toolbar to reflect these changes
    redrawToolbar

End Sub

'Given mouse coordinates over the form, return the thumbnail at that location.  If the cursor is not over a thumbnail,
' the function will return -1
Private Function getThumbAtPosition(ByVal x As Long, ByVal y As Long) As Long
    
    Dim thumbOffset As Long
    thumbOffset = hsThumbnails.Value
    
    If verticalLayout Then
        getThumbAtPosition = (y + thumbOffset) \ fixDPI(thumbHeight)
        If getThumbAtPosition > (numOfThumbnails - 1) Then getThumbAtPosition = -1
    Else
        getThumbAtPosition = (x + thumbOffset) \ fixDPI(thumbWidth)
        If getThumbAtPosition > (numOfThumbnails - 1) Then getThumbAtPosition = -1
    End If
    
End Function

Public Sub cMouseEvents_MouseHScroll(ByVal CharsScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)

    'Horizontal scrolling - only trigger it if the horizontal scroll bar is actually visible
    If m_ListScrollable Then
  
        If CharsScrolled < 0 Then
            
            If hsThumbnails.Value + hsThumbnails.LargeChange > hsThumbnails.Max Then
                hsThumbnails.Value = hsThumbnails.Max
            Else
                hsThumbnails.Value = hsThumbnails.Value + hsThumbnails.LargeChange
            End If
            
            curThumbHover = getThumbAtPosition(x, y)
            redrawToolbar
        
        ElseIf CharsScrolled > 0 Then
            
            If hsThumbnails.Value - hsThumbnails.LargeChange < hsThumbnails.Min Then
                hsThumbnails.Value = hsThumbnails.Min
            Else
                hsThumbnails.Value = hsThumbnails.Value - hsThumbnails.LargeChange
            End If
            
            curThumbHover = getThumbAtPosition(x, y)
            redrawToolbar
            
        End If
        
    End If
    
End Sub

Private Sub cMouseEvents_MouseIn()
    g_MouseOverImageTabstrip = True
End Sub

Private Sub cMouseEvents_MouseOut()
        
    g_MouseOverImageTabstrip = False
    
    If curThumbHover <> -1 Then
        curThumbHover = -1
        redrawToolbar
    End If
    
    cMouseEvents.MousePointer = 0
    
End Sub

Public Sub cMouseEvents_MouseVScroll(ByVal LinesScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)

    'Vertical scrolling - only trigger it if the horizontal scroll bar is actually visible
    If m_ListScrollable Then
  
        If LinesScrolled < 0 Then
            
            If hsThumbnails.Value + hsThumbnails.LargeChange > hsThumbnails.Max Then
                hsThumbnails.Value = hsThumbnails.Max
            Else
                hsThumbnails.Value = hsThumbnails.Value + hsThumbnails.LargeChange
            End If
            
            curThumbHover = getThumbAtPosition(x, y)
            redrawToolbar
        
        ElseIf LinesScrolled > 0 Then
            
            If hsThumbnails.Value - hsThumbnails.LargeChange < hsThumbnails.Min Then
                hsThumbnails.Value = hsThumbnails.Min
            Else
                hsThumbnails.Value = hsThumbnails.Value - hsThumbnails.LargeChange
            End If
            
            curThumbHover = getThumbAtPosition(x, y)
            redrawToolbar
            
        End If
        
    End If

End Sub

Private Sub Form_Load()

    'Reset the thumbnail array
    numOfThumbnails = 0
    ReDim imgThumbnails(0 To numOfThumbnails) As thumbEntry
    
    'Enable mousewheel scrolling
    Set cMouseEvents = New bluMouseEvents
    cMouseEvents.Attach Me.hWnd
    cMouseEvents.MousePointer = IDC_HAND
    
    'Detect initial alignment
    If (g_WindowManager.getImageTabstripAlignment = vbAlignLeft) Or (g_WindowManager.getImageTabstripAlignment = vbAlignRight) Then
        verticalLayout = True
    Else
        verticalLayout = False
    End If
    
    'Set default thumbnail sizes
    If verticalLayout Then
        thumbWidth = g_WindowManager.getClientWidth(Me.hWnd)
        thumbHeight = thumbWidth
    Else
        thumbHeight = g_WindowManager.getClientHeight(Me.hWnd)
        thumbWidth = thumbHeight
    End If
    
    'Retrieve the unsaved image notification icon from the resource file
    Set unsavedChangesLayer = New pdLayer
    loadResourceToLayer "NTFY_UNSAVED", unsavedChangesLayer
    
    'Update the drop-shadow blur radius to account for DPI
    shadowBlurRadius = fixDPI(2)
    
    'If the tabstrip ever becomes long enough to scroll, this will be set to TRUE
    m_ListScrollable = False
    
    'Activate the custom tooltip handler
    Set m_ToolTip = New clsToolTip
    m_ToolTip.Create Me
    m_ToolTip.MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
    m_ToolTip.DelayTime(ttDelayShow) = 10000
    m_ToolTip.AddTool Me, ""
    
    'Theme the form
    makeFormPretty Me
    
End Sub

'When the left mouse button is pressed, activate click-to-drag mode for scrolling the tabstrip window
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Make a note of the initial mouse position
    If Button = vbLeftButton Then
        m_MouseDown = True
        m_InitX = x
        m_InitY = y
        m_MouseDistanceTraveled = 0
        m_InitOffset = hsThumbnails.Value
    End If
    
    'Reset the "resize in progress" tracker
    weAreResponsibleForResize = False
    
    'Reset the "scrolling occured" tracker
    m_ScrollingOccured = False
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Note that the mouse is currently over the tabstrip
    g_MouseOverImageTabstrip = True
    
    'We require a few mouse movements to fire before doing anything; otherwise this function will fire constantly.
    m_MouseDistanceTraveled = m_MouseDistanceTraveled + 1
    
    'We handle several different _MouseMove scenarios, in this order:
    ' 1) If the mouse is near the resizable edge of the form, and the left button is depressed, activate live resizing.
    ' 2) If a button is depressed, activate tabstrip scrolling (if the list is long enough)
    ' 3) If no buttons are depressed, hover the image at the current position (if any)
    
    'If the mouse is near the resizable edge of the toolbar (which varies according to its alignment),
    ' allow the user to resize the thumbnail toolbar
    Dim mouseInResizeTerritory As Boolean
    
    'How close does the mouse have to be to the form border to allow resizing; currently we use 7 pixels, while accounting
    ' for DPI variance (e.g. 7 pixels at 96 dpi)
    Dim resizeBorderAllowance As Long
    resizeBorderAllowance = fixDPI(7)
    
    Dim hitCode As Long
    
    Select Case g_WindowManager.getImageTabstripAlignment
    
        Case vbAlignLeft
            If (y > 0) And (y < Me.ScaleHeight) And (x > Me.ScaleWidth - resizeBorderAllowance) Then mouseInResizeTerritory = True
            hitCode = HTRIGHT
        
        Case vbAlignTop
            If (x > 0) And (x < Me.ScaleWidth) And (y > Me.ScaleHeight - resizeBorderAllowance) Then mouseInResizeTerritory = True
            hitCode = HTBOTTOM
        
        Case vbAlignRight
            If (y > 0) And (y < Me.ScaleHeight) And (x < resizeBorderAllowance) Then mouseInResizeTerritory = True
            hitCode = HTLEFT
        
        Case vbAlignBottom
            If (x > 0) And (x < Me.ScaleWidth) And (y < resizeBorderAllowance) Then mouseInResizeTerritory = True
            hitCode = HTTOP
    
    End Select
        
    'Check mouse button state; if it's down, check for resize or scrolling of the image list
    If m_MouseDown Then
        
        If mouseInResizeTerritory Then
                
            If Button = vbLeftButton Then
                
                'Allow resizing
                weAreResponsibleForResize = True
                ReleaseCapture
                SendMessage Me.hWnd, WM_NCLBUTTONDOWN, hitCode, ByVal 0&
                
            End If
        
        'The mouse is not in resize territory.
        Else
        
            mouseInResizeTerritory = False
            
            'If the list is scrollable (due to tons of images being loaded), calculate a new offset now
            If m_ListScrollable And (m_MouseDistanceTraveled > 5) And (Not weAreResponsibleForResize) Then
            
                m_ScrollingOccured = True
            
                Dim mouseOffset As Long
                
                If verticalLayout Then
                    mouseOffset = (m_InitY - y)
                Else
                    mouseOffset = (m_InitX - x)
                End If
                
                'Change the invisible scroll bar to match the new offset
                Dim newScrollValue As Long
                newScrollValue = m_InitOffset + mouseOffset
                
                If newScrollValue < 0 Then
                    hsThumbnails.Value = 0
                
                ElseIf newScrollValue > hsThumbnails.Max Then
                    hsThumbnails.Value = hsThumbnails.Max
                    
                Else
                    hsThumbnails.Value = newScrollValue
                    
                End If
                
            
            End If
        
        End If
    
    'The left mouse button is not down.  Hover the image beneath the cursor (if any)
    Else
    
        Dim oldThumbHover As Long
        oldThumbHover = curThumbHover
        
        'Retrieve the thumbnail at this position, and change the mouse pointer accordingly
        curThumbHover = getThumbAtPosition(x, y)
        
        'To prevent flickering, only update the tooltip when absolutely necessary
        If curThumbHover <> oldThumbHover Then
        
            'If the cursor is over a thumbnail, update the tooltip to display that image's filename
            If curThumbHover <> -1 Then
                        
                If Len(pdImages(imgThumbnails(curThumbHover).indexInPDImages).locationOnDisk) > 0 Then
                    m_ToolTip.ToolTipHeader = pdImages(imgThumbnails(curThumbHover).indexInPDImages).originalFileNameAndExtension
                    m_ToolTip.ToolText(Me) = pdImages(imgThumbnails(curThumbHover).indexInPDImages).locationOnDisk
                Else
                    m_ToolTip.ToolTipHeader = g_Language.TranslateMessage("This image does not have a filename.")
                    m_ToolTip.ToolText(Me) = g_Language.TranslateMessage("Once this image has been saved to disk, its filename will appear here.")
                End If
            
            'The cursor is not over a thumbnail; let the user know they can hover if they want more information.
            Else
            
                m_ToolTip.ToolTipHeader = ""
                m_ToolTip.ToolText(Me) = "Hover an image thumbnail to see its name and current file location."
            
            End If
            
        End If
        
    End If
    
    'Set a mouse pointer according to the handling above
    If mouseInResizeTerritory Then
    
        If verticalLayout Then
            cMouseEvents.MousePointer = IDC_SIZEWE
        Else
            cMouseEvents.MousePointer = IDC_SIZENS
        End If
            
    Else
    
        'Display a hand cursor if over an image
        If curThumbHover = -1 Then cMouseEvents.MousePointer = vbDefault Else cMouseEvents.MousePointer = IDC_HAND
    
    End If
    
    'Regardless of what happened above, redraw the toolbar to reflect any changes
    redrawToolbar
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If the _MouseUp event was triggered by the user, select the image at that position
    If Not weAreResponsibleForResize Then
    
        Dim potentialNewThumb As Long
        potentialNewThumb = getThumbAtPosition(x, y)
        
        'Notify the program that a new image has been selected; it will then bring that image to the foreground,
        ' which will automatically trigger a toolbar redraw.  Also, do not select the image if the user has been
        ' scrolling the list.
        If (potentialNewThumb >= 0) And (Not m_ScrollingOccured) Then
            curThumb = potentialNewThumb
            pdImages(imgThumbnails(curThumb).indexInPDImages).containingForm.ActivateWorkaround "user clicked image thumbnail"
        End If
        
    End If
    
    'Release mouse tracking
    If m_MouseDown Then
        m_MouseDown = False
        m_InitX = 0
        m_InitY = 0
        m_MouseDistanceTraveled = 0
    End If

End Sub

'Any time this window is resized, we need to recreate the thumbnail display
Private Sub Form_Resize()

    'Detect alignment changes (if any)
    If (g_WindowManager.getImageTabstripAlignment = vbAlignLeft) Or (g_WindowManager.getImageTabstripAlignment = vbAlignRight) Then
        verticalLayout = True
    Else
        verticalLayout = False
    End If

    Dim i As Long

    'If the tabstrip is horizontal and the window's height is changing, we need to recreate all image thumbnails
    If ((Not verticalLayout) And (thumbHeight <> g_WindowManager.getClientHeight(Me.hWnd))) Then
        
        thumbHeight = g_WindowManager.getClientHeight(Me.hWnd)
    
        For i = 0 To numOfThumbnails - 1
            imgThumbnails(i).thumbLayer.eraseLayer
            pdImages(imgThumbnails(i).indexInPDImages).requestThumbnail imgThumbnails(i).thumbLayer, fixDPI(thumbHeight) - (fixDPI(thumbBorder) * 2)
            updateShadowLayer i
        Next i
    
    End If
    
    'If the tabstrip is vertical and the window's with is changing, we need to recreate all image thumbnails
    If (verticalLayout And (thumbWidth <> g_WindowManager.getClientWidth(Me.hWnd))) Then
    
        thumbWidth = g_WindowManager.getClientWidth(Me.hWnd)
        
        For i = 0 To numOfThumbnails - 1
            imgThumbnails(i).thumbLayer.eraseLayer
            pdImages(imgThumbnails(i).indexInPDImages).requestThumbnail imgThumbnails(i).thumbLayer, fixDPI(thumbWidth) - (fixDPI(thumbBorder) * 2)
            updateShadowLayer i
        Next i
    
    End If
    
    'Update thumbnail sizes
    If verticalLayout Then
        thumbWidth = g_WindowManager.getClientWidth(Me.hWnd)
        thumbHeight = thumbWidth
    Else
        thumbHeight = g_WindowManager.getClientHeight(Me.hWnd)
        thumbWidth = thumbHeight
    End If
        
    'Create a background buffer the same size as this window
    m_BufferWidth = g_WindowManager.getClientWidth(Me.hWnd)
    m_BufferHeight = g_WindowManager.getClientHeight(Me.hWnd)
    
    'Redraw the toolbar
    redrawToolbar
    
    'Notify the window manager that the tab strip has been resized; it will resize image windows to match
    'If Not weAreResponsibleForResize Then
    g_WindowManager.notifyImageTabStripResized
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    g_WindowManager.unregisterForm Me
    Set cMouseEvents = Nothing
End Sub

'Whenever a function wants to render the current toolbar, it may do so here
Private Sub redrawToolbar()

    'Recreate the toolbar buffer
    Set bufferLayer = New pdLayer
    bufferLayer.createBlank m_BufferWidth, m_BufferHeight, 24, ConvertSystemColor(vb3DShadow)
    
    'Horizontal/vertical layout changes the constraining dimension (e.g. the dimension used to detect if the number
    ' of image tabs currently visible is long enough that it needs to be scrollable).
    Dim constrainingDimension As Long, constrainingMax As Long
    If verticalLayout Then
        constrainingDimension = thumbHeight
        constrainingMax = m_BufferHeight
    Else
        constrainingDimension = thumbWidth
        constrainingMax = m_BufferWidth
    End If
    
    'Determine if the scrollbar needs to be accounted for or not
    Dim maxThumbSize As Long
    maxThumbSize = fixDPIFloat(constrainingDimension) * numOfThumbnails - 1
    
    If maxThumbSize < constrainingMax Then
        hsThumbnails.Value = 0
        m_ListScrollable = False
    Else
        m_ListScrollable = True
        hsThumbnails.Max = maxThumbSize - constrainingMax
        
        'Dynamically set the scrollbar's LargeChange value relevant to thumbnail size
        Dim lChange As Long
        
        lChange = (maxThumbSize - constrainingMax) \ 16
        
        If lChange < 1 Then lChange = 1
        If lChange > thumbWidth \ 4 Then lChange = thumbWidth \ 4
        
        hsThumbnails.LargeChange = lChange
        
    End If
    
    'Determine a scrollbar offset as necessary
    Dim scrollOffset As Long
    scrollOffset = hsThumbnails.Value
    
    'Render each thumbnail block
    Dim i As Long
    For i = 0 To numOfThumbnails - 1
        If verticalLayout Then
            renderThumbTab i, 0, fixDPI(i * thumbHeight) - scrollOffset
        Else
            renderThumbTab i, fixDPI(i * thumbWidth) - scrollOffset, 0
        End If
    Next i
    
    'Eventually we'll do something nicer, but for now, draw a line across the edge of the tabstrip nearest the image.
    Select Case g_WindowManager.getImageTabstripAlignment
    
        Case vbAlignLeft
            GDIPlusDrawLineToDC bufferLayer.getLayerDC, m_BufferWidth - 1, 0, m_BufferWidth - 1, m_BufferHeight, ConvertSystemColor(vb3DLight), 255, 2, False
        
        Case vbAlignTop
            GDIPlusDrawLineToDC bufferLayer.getLayerDC, 0, m_BufferHeight - 1, m_BufferWidth, m_BufferHeight - 1, ConvertSystemColor(vb3DLight), 255, 2, False
        
        Case vbAlignRight
            GDIPlusDrawLineToDC bufferLayer.getLayerDC, 1, 0, 1, m_BufferHeight, ConvertSystemColor(vb3DLight), 255, 2, False
        
        Case vbAlignBottom
            GDIPlusDrawLineToDC bufferLayer.getLayerDC, 0, 1, m_BufferWidth, 1, ConvertSystemColor(vb3DLight), 255, 2, False
    
    End Select
    
    'Activate color management for our form
    assignDefaultColorProfileToObject Me.hWnd, Me.hDC
    turnOnColorManagementForDC Me.hDC
    
    'Copy the buffer to the form
    BitBlt Me.hDC, 0, 0, m_BufferWidth, m_BufferHeight, bufferLayer.getLayerDC, 0, 0, vbSrcCopy
    Me.Picture = Me.Image
    Me.Refresh
    
End Sub
    
'Render a given thumbnail onto the background form at the specified offset
Private Sub renderThumbTab(ByVal thumbIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current tab if it will be visible
    Dim tabVisible As Boolean
    tabVisible = False
    
    If verticalLayout Then
        If ((offsetY + fixDPI(thumbHeight)) > 0) And (offsetY < m_BufferHeight) Then tabVisible = True
    Else
        If ((offsetX + fixDPI(thumbWidth)) > 0) And (offsetX < m_BufferWidth) Then tabVisible = True
    End If
    
    If tabVisible Then
    
        Dim tmpRect As RECT
        Dim hBrush As Long
    
        'If this thumbnail has been selected, draw the background with the system's current selection color
        If thumbIndex = curThumb Then
            SetRect tmpRect, offsetX, offsetY, offsetX + fixDPI(thumbWidth), offsetY + fixDPI(thumbHeight)
            hBrush = CreateSolidBrush(ConvertSystemColor(vb3DLight))
            FillRect bufferLayer.getLayerDC, tmpRect, hBrush
            DeleteObject hBrush
        End If
        
        'If the current thumbnail is highlighted but not selected, simply render the border with a highlight
        If (thumbIndex <> curThumb) And (thumbIndex = curThumbHover) Then
            SetRect tmpRect, offsetX, offsetY, offsetX + fixDPI(thumbWidth), offsetY + fixDPI(thumbHeight)
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FrameRect bufferLayer.getLayerDC, tmpRect, hBrush
            SetRect tmpRect, tmpRect.Left + 1, tmpRect.Top + 1, tmpRect.Right - 1, tmpRect.Bottom - 1
            FrameRect bufferLayer.getLayerDC, tmpRect, hBrush
            DeleteObject hBrush
        End If
    
        'Render the matching thumbnail shadow and thumbnail into this block
        imgThumbnails(thumbIndex).thumbShadow.alphaBlendToDC bufferLayer.getLayerDC, 192, offsetX, offsetY + fixDPI(1)
        imgThumbnails(thumbIndex).thumbLayer.alphaBlendToDC bufferLayer.getLayerDC, 255, offsetX + fixDPI(thumbBorder), offsetY + fixDPI(thumbBorder)
        
        'If the parent image has unsaved changes, also render a notification icon
        If Not pdImages(imgThumbnails(thumbIndex).indexInPDImages).getSaveState Then
            unsavedChangesLayer.alphaBlendToDC bufferLayer.getLayerDC, 230, offsetX + fixDPI(thumbBorder) + fixDPI(2), offsetY + fixDPI(thumbHeight) - fixDPI(thumbBorder) - unsavedChangesLayer.getLayerHeight - fixDPI(2)
        End If
        
    End If

End Sub

'Whenever a thumbnail has been updated, this sub must be called to regenerate its drop-shadow
Private Sub updateShadowLayer(ByVal imgThumbnailIndex As Long)
    imgThumbnails(imgThumbnailIndex).thumbShadow.eraseLayer
    createShadowLayer imgThumbnails(imgThumbnailIndex).thumbLayer, imgThumbnails(imgThumbnailIndex).thumbShadow
    padLayer imgThumbnails(imgThumbnailIndex).thumbShadow, fixDPI(thumbBorder)
    quickBlurLayer imgThumbnails(imgThumbnailIndex).thumbShadow, shadowBlurRadius
    imgThumbnails(imgThumbnailIndex).thumbShadow.fixPremultipliedAlpha True
End Sub

'Even though the scroll bar is not visible, we still process mousewheel events using it, so redraw when it changes
Private Sub hsThumbnails_Change()
    redrawToolbar
End Sub

Private Sub hsThumbnails_Scroll()
    redrawToolbar
End Sub

'External functions can use this to re-theme this form at run-time (important when changing languages, for example)
Public Sub requestMakeFormPretty()
    makeFormPretty Me, m_ToolTip
End Sub
