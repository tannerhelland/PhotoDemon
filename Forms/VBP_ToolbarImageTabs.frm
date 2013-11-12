VERSION 5.00
Begin VB.Form toolbar_ImageTabs 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
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
'Copyright ©2012-2013 by Tanner Helland
'Created: 15/October/13
'Last updated: 21/October/13
'Last update: fix syncing of curThumb and g_CurrentImage when an inactive image is unloaded via the taskbar
'
'In fall 2013, PhotoDemon left behind the MDI model in favor of fully dockable/floatable tool and image windows.
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
            pdImages(pdImagesIndex).requestThumbnail imgThumbnails(i).thumbLayer, fixDPI(thumbWidth) - (fixDPI(thumbBorder) * 2)
            updateShadowLayer i
            imgThumbnails(i).thumbLayer.fixPremultipliedAlpha True
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
    pdImages(pdImagesIndex).requestThumbnail imgThumbnails(numOfThumbnails).thumbLayer, fixDPI(thumbWidth) - (fixDPI(thumbBorder) * 2)
    
    'Create a matching shadow layer
    Set imgThumbnails(numOfThumbnails).thumbShadow = New pdLayer
    updateShadowLayer numOfThumbnails
    
    imgThumbnails(numOfThumbnails).thumbLayer.fixPremultipliedAlpha True
    
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
    
    Dim hOffset As Long
    hOffset = hsThumbnails.Value
    
    getThumbAtPosition = (x + hOffset) \ fixDPI(thumbWidth)
    If getThumbAtPosition > (numOfThumbnails - 1) Then getThumbAtPosition = -1
    
End Function

Private Sub cMouseEvents_MouseHScroll(ByVal CharsScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)

    'Horizontal scrolling - only trigger it if the horizontal scroll bar is actually visible
    If hsThumbnails.Visible Then
  
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

Private Sub cMouseEvents_MouseOut()
    If curThumbHover <> -1 Then
        curThumbHover = -1
        redrawToolbar
    End If
    cMouseEvents.MousePointer = 0
End Sub

Private Sub cMouseEvents_MouseVScroll(ByVal LinesScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)

    'Vertical scrolling - only trigger it if the horizontal scroll bar is actually visible
    If hsThumbnails.Visible Then
  
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
    
    'Set default thumbnail sizes
    thumbHeight = g_WindowManager.getClientHeight(Me.hWnd)
    thumbWidth = thumbHeight
    
    'Retrieve the unsaved image notification icon from the resource file
    Set unsavedChangesLayer = New pdLayer
    loadResourceToLayer "NTFY_UNSAVED", unsavedChangesLayer
    
    'Update the drop-shadow blur radius to account for DPI
    shadowBlurRadius = fixDPI(2)
    
    'Activate the custom tooltip handler
    Set m_ToolTip = New clsToolTip
    m_ToolTip.Create Me
    m_ToolTip.MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
    m_ToolTip.DelayTime(ttDelayShow) = 10000
    m_ToolTip.AddTool Me, ""
    
    'Theme the form
    makeFormPretty Me
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Not weAreResponsibleForResize Then
    
        Dim potentialNewThumb As Long
        potentialNewThumb = getThumbAtPosition(x, y)
        
        'Notify the program that a new image has been selected; it will then bring that image to the foreground,
        ' which will automatically trigger a toolbar redraw
        If potentialNewThumb >= 0 Then
            curThumb = potentialNewThumb
            pdImages(imgThumbnails(curThumb).indexInPDImages).containingForm.ActivateWorkaround "user clicked image thumbnail"
        End If
        
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
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
        
        Else
        
            m_ToolTip.ToolTipHeader = ""
            m_ToolTip.ToolText(Me) = "Hover an image thumbnail to see its name and current file location."
        
        End If
        
    End If
    
    'If the mouse is near the bottom of the toolbar, allow the user to resize the thumbnail toolbar
    If (x > 0) And (x < Me.ScaleWidth) And (y > Me.ScaleHeight - 6) Then
        cMouseEvents.MousePointer = IDC_SIZENS
        
        If Button = vbLeftButton Then
            
            'Allow resizing
            weAreResponsibleForResize = True
            ReleaseCapture
            SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOM, ByVal 0&
            
        End If
        
    Else
        weAreResponsibleForResize = False
        If curThumbHover = -1 Then cMouseEvents.MousePointer = vbDefault Else cMouseEvents.MousePointer = IDC_HAND
    End If
    
    redrawToolbar
    
End Sub

'Any time this window is resized, we need to recreate the thumbnail display
Private Sub Form_Resize()

    'If the window's height is changing, we need to redraw all image thumbnails
    If thumbHeight <> g_WindowManager.getClientHeight(Me.hWnd) Then
        
        thumbHeight = g_WindowManager.getClientHeight(Me.hWnd)
    
        Dim i As Long
        For i = 0 To numOfThumbnails - 1
            imgThumbnails(i).thumbLayer.eraseLayer
            pdImages(imgThumbnails(i).indexInPDImages).requestThumbnail imgThumbnails(i).thumbLayer, fixDPI(thumbHeight) - (fixDPI(thumbBorder) * 2)
            updateShadowLayer i
            imgThumbnails(i).thumbLayer.fixPremultipliedAlpha True
        Next i
    
    End If
    
    'Update thumbnail sizes
    thumbHeight = g_WindowManager.getClientHeight(Me.hWnd)
    thumbWidth = thumbHeight
    
    'Create a background buffer the same size as this window
    m_BufferWidth = g_WindowManager.getClientWidth(Me.hWnd)
    m_BufferHeight = g_WindowManager.getClientHeight(Me.hWnd)
    
    'Redraw the toolbar
    redrawToolbar
    
    'Resize and position the scroll bar
    hsThumbnails.Move 0, m_BufferHeight - hsThumbnails.Height - 2, m_BufferWidth, hsThumbnails.Height
    
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
    
    'Determine if the horizontal scrollbar needs to be visible or not
    Dim maxThumbSize As Long
    maxThumbSize = fixDPIFloat(thumbWidth) * numOfThumbnails - 1
    
    If maxThumbSize < m_BufferWidth Then
        hsThumbnails.Value = 0
        If hsThumbnails.Visible Then hsThumbnails.Visible = False
    Else
        If Not hsThumbnails.Visible Then hsThumbnails.Visible = True
        hsThumbnails.Max = maxThumbSize - m_BufferWidth
    End If
    
    'Determine a scrollbar offset as necessary
    Dim scrollOffset As Long
    scrollOffset = hsThumbnails.Value
    
    'Render each thumbnail block
    Dim i As Long
    For i = 0 To numOfThumbnails - 1
        renderThumbTab i, fixDPI(i * thumbWidth) - scrollOffset, 0
    Next i
    
    'Eventually we'll do something nicer, but for now, draw a line across the bottom of the tabstrip
    GDIPlusDrawLineToDC bufferLayer.getLayerDC, 0, m_BufferHeight - 1, m_BufferWidth, m_BufferHeight - 1, ConvertSystemColor(vb3DLight), 255, 2, False
    
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
    If ((offsetX + fixDPI(thumbWidth)) > 0) And (offsetX < m_BufferWidth) Then
    
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

Private Sub hsThumbnails_Change()
    redrawToolbar
End Sub

Private Sub hsThumbnails_Scroll()
    redrawToolbar
End Sub

'Whenever a thumbnail has been updated, this sub must be called to regenerate its drop-shadow
Private Sub updateShadowLayer(ByVal imgThumbnailIndex As Long)
    imgThumbnails(imgThumbnailIndex).thumbShadow.eraseLayer
    createShadowLayer imgThumbnails(imgThumbnailIndex).thumbLayer, imgThumbnails(imgThumbnailIndex).thumbShadow
    padLayer imgThumbnails(imgThumbnailIndex).thumbShadow, fixDPI(thumbBorder)
    quickBlurLayer imgThumbnails(imgThumbnailIndex).thumbShadow, shadowBlurRadius
    imgThumbnails(imgThumbnailIndex).thumbShadow.fixPremultipliedAlpha True
End Sub
