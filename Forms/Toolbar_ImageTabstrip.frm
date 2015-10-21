VERSION 5.00
Begin VB.Form toolbar_ImageTabs 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Images"
   ClientHeight    =   1140
   ClientLeft      =   2250
   ClientTop       =   1770
   ClientWidth     =   13725
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
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   915
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar hsThumbnails 
      Height          =   255
      Left            =   0
      Max             =   10
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   13695
   End
   Begin VB.Menu mnuImageTabsContext 
      Caption         =   "&Image"
      Visible         =   0   'False
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "Save copy (&lossless)"
         Index           =   1
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "Save &as..."
         Index           =   2
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "Revert"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "Open location in E&xplorer"
         Index           =   5
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "&Close"
         Index           =   7
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "Close all except this"
         Index           =   8
      End
   End
End
Attribute VB_Name = "toolbar_ImageTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Selection ("Tab") Toolbar
'Copyright 2013-2015 by Tanner Helland
'Created: 15/October/13
'Last updated: 19/February/15
'Last updated by: Raj
'Last update: Added a close icon on hover of each thumbnail, and a context menu
'
'In fall 2013, PhotoDemon left behind the MDI model in favor of fully dockable/floatable tool and image windows.
' This required quite a new features, including a way to switch between loaded images when image windows are docked -
' which is where this form comes in.
'
'The purpose of this form is to provide a tab-like interface for switching between open images.  Please note that
' much of this form's layout and alignment is handled by PhotoDemon's window manager, so you will need to look
' there for detailed information on things like the window's positioning and alignment.
'
'To my knowledge, as of January '14 the tabstrip should work properly under all orientations and screen DPIs.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'A collection of all currently active thumbnails; this is dynamically resized as thumbnails are added/removed.
Private Type thumbEntry
    thumbDIB As pdDIB
    thumbShadow As pdDIB
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
Private bufferDIB As pdDIB
Private m_BufferWidth As Long, m_BufferHeight As Long

'An outside class provides access to mousewheel events for scrolling the tabstrip view
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'The currently selected and currently hovered thumbnail
Private curThumb As Long, curThumbHover As Long

'When we are responsible for this window resizing (because the user is resizing our window manually), we set this to TRUE.
' This variable is then checked before requesting additional redraws during our resize event.
Private weAreResponsibleForResize As Boolean

'As a convenience to the user, we provide a small notification when an image has unsaved changes
Private unsavedChangesDIB As pdDIB

'In Feb '15, Raj added the very nice capability to close an image by hovering its tab, then clicking the X that magically appears.
' A few DIBs are required for this: normal gray, red highlight when hovered, and an underlying shadow (to help it stand out against
' dark thumbnails).
Private m_CloseIconRed As pdDIB, m_CloseIconGray As pdDIB, m_CloseIconShadow As pdDIB

'We also need a few tracking variables, for example if the user closes a tab that is *not* currently the active one
Private m_CloseTriggeredOnThumbnail As Long
Private m_CloseIconHovered As Long

'Thumbnails can be right-clicked to see a context menu
Private m_RightClickedThumbnail As Long

'Drop-shadows on the thumbnails have a variable radius that changes based on the user's DPI settings
Private shadowBlurRadius As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Private toolTipManager As pdToolTip

'If the user loads tons of images, the tabstrip may overflow the available area.  We now allow them to drag-scroll the list.
' In order to allow that, we must track a few extra things, like initial mouse x/y.
Private m_MouseDown As Boolean, m_ScrollingOccured As Boolean
Private m_InitX As Long, m_InitY As Long, m_InitOffset As Long
Private m_ListScrollable As Boolean
Private m_MouseDistanceTraveled As Long

'Most importantly for scrolling, this value is set to TRUE on cMouseEvents_MouseDownCustom, *if* the mouse is clicked near the resizable edge of the
' toolbar (which varies according to its alignment, obviously).
Private m_MouseInResizeTerritory As Boolean

'Horizontal or vertical layout; obviously, all our rendering and mouse detection code changes depending on the orientation
' of the tabstrip.
Private verticalLayout As Boolean

'In Feb '15, Raj added a great context menu to the tabstrip.  To help simplify menu enable/disable behavior, this enum can be used to identify
' individual menu entries.
Private Enum POPUP_MENU_ENTRIES
    POP_SAVE = 0
    POP_SAVE_COPY = 1
    POP_SAVE_AS = 2
    POP_REVERT = 3
    POP_OPEN_IN_EXPLORER = 5
    POP_CLOSE = 7
    POP_CLOSE_OTHERS = 8
End Enum

#If False Then
    Private Const POP_SAVE = 0, POP_SAVE_COPY = 1, POP_SAVE_AS = 2, POP_REVERT = 3, POP_OPEN_IN_EXPLORER = 5, POP_CLOSE = 7, POP_CLOSE_OTHERS = 8
#End If

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
        Else
            curThumb = 0
        End If
    Next i
    
    'Redraw the toolbar to reflect the change
    redrawToolbar True
        
End Sub

'Returns TRUE is a given thumbnail is currently viewable in its entirety; FALSE if it lies partially or fully off-screen.
Public Function fitThumbnailOnscreen(ByVal thumbIndex As Long) As Boolean

    Dim isThumbnailOnscreen As Boolean

    'First, figure out where the thumbnail actually sits.
    
    'Determine a scrollbar offset as necessary
    Dim scrollOffset As Long
    scrollOffset = hsThumbnails.Value
    
    'Per the tabstrip's current alignment, figure out a relevant position
    Dim hPosition As Long, vPosition As Long
    
    If verticalLayout Then
        hPosition = 0
        vPosition = (thumbIndex * thumbHeight) - scrollOffset
    Else
        hPosition = (thumbIndex * thumbWidth) - scrollOffset
        vPosition = 0
    End If
    
    'Use the tabstrip's size to determine if this thumbnail lies off-screen
    If verticalLayout Then
        
        If vPosition < 0 Or (vPosition + thumbHeight - 1) > Me.ScaleHeight Then
            isThumbnailOnscreen = False
        Else
            isThumbnailOnscreen = True
        End If
        
    Else
    
        If hPosition < 0 Or (hPosition + thumbWidth - 1) > Me.ScaleWidth Then
            isThumbnailOnscreen = False
        Else
            isThumbnailOnscreen = True
        End If
        
    End If
    
    'If the thumbnail is not onscreen, make it so!
    If Not isThumbnailOnscreen Then
    
        If verticalLayout Then
        
            If vPosition < 0 Then
                hsThumbnails.Value = thumbIndex * thumbHeight
            Else
            
                If ((thumbIndex + 1) * thumbHeight) - Me.ScaleHeight > hsThumbnails.Max Then
                    hsThumbnails.Value = hsThumbnails.Max
                Else
                    hsThumbnails.Value = ((thumbIndex + 1) * thumbHeight) - Me.ScaleHeight
                End If
                
            End If
            
        Else
        
            If hPosition < 0 Then
                hsThumbnails.Value = thumbIndex * thumbWidth
            Else
            
                If ((thumbIndex + 1) * thumbWidth) - Me.ScaleWidth > hsThumbnails.Max Then
                    hsThumbnails.Value = hsThumbnails.Max
                Else
                    hsThumbnails.Value = ((thumbIndex + 1) * thumbWidth) - Me.ScaleWidth
                End If
                
            End If
            
        End If
    
    End If
            
End Function

'When the user somehow changes an image, they need to notify the toolbar, so that a new thumbnail can be rendered
Public Sub notifyUpdatedImage(ByVal pdImagesIndex As Long)
    
    'Find the matching thumbnail entry, and update its thumbnail DIB
    Dim i As Long
    For i = 0 To numOfThumbnails
        If imgThumbnails(i).indexInPDImages = pdImagesIndex Then
            
            If Not (pdImages(pdImagesIndex) Is Nothing) Then
            
                If verticalLayout Then
                    pdImages(pdImagesIndex).requestThumbnail imgThumbnails(i).thumbDIB, thumbHeight - (FixDPI(thumbBorder) * 2)
                Else
                    pdImages(pdImagesIndex).requestThumbnail imgThumbnails(i).thumbDIB, thumbWidth - (FixDPI(thumbBorder) * 2)
                End If
                
            End If
            
            If g_InterfacePerformance <> PD_PERF_FASTEST Then updateShadowDIB i
            Exit For
        End If
    Next i
    
    'Redraw the toolbar to reflect the change
    redrawToolbar
        
End Sub

'Whenever a new image is loaded, it needs to be registered with the toolbar
Public Sub registerNewImage(ByVal pdImagesIndex As Long)

    'Request a thumbnail from the relevant pdImage object, and premultiply it to allow us to blit it more quickly
    Set imgThumbnails(numOfThumbnails).thumbDIB = New pdDIB
    
    If verticalLayout Then
        pdImages(pdImagesIndex).requestThumbnail imgThumbnails(numOfThumbnails).thumbDIB, thumbHeight - (FixDPI(thumbBorder) * 2)
    Else
        pdImages(pdImagesIndex).requestThumbnail imgThumbnails(numOfThumbnails).thumbDIB, thumbWidth - (FixDPI(thumbBorder) * 2)
    End If
    
    'Create a matching shadow DIB
    Set imgThumbnails(numOfThumbnails).thumbShadow = New pdDIB
    If g_InterfacePerformance <> PD_PERF_FASTEST Then updateShadowDIB numOfThumbnails
    
    'Make a note of this thumbnail's index in the main pdImages array
    imgThumbnails(numOfThumbnails).indexInPDImages = pdImagesIndex
    
    'We can assume this image will be the active one
    curThumb = numOfThumbnails
    
    'Prepare the array to receive another entry in the future
    numOfThumbnails = numOfThumbnails + 1
    ReDim Preserve imgThumbnails(0 To numOfThumbnails) As thumbEntry
    
    'Redraw the toolbar to reflect these changes
    redrawToolbar True
    
End Sub

'Whenever an image is unloaded, it needs to be de-registered with the toolbar
Public Sub RemoveImage(ByVal pdImagesIndex As Long, Optional ByVal refreshToolbar As Boolean = True)

    'Find the matching thumbnail in our collection
    Dim i As Long, thumbIndex As Long
    thumbIndex = -1
    
    For i = 0 To numOfThumbnails
        If imgThumbnails(i).indexInPDImages = pdImagesIndex Then
            thumbIndex = i
            Exit For
        End If
    Next i
    
    'thumbIndex is now equal to the matching thumbnail.  Remove that entry, then shift all thumbnails after that point down.
    If (thumbIndex > -1) And (thumbIndex <= UBound(imgThumbnails)) Then
    
        If Not (imgThumbnails(thumbIndex).thumbDIB Is Nothing) Then
            imgThumbnails(thumbIndex).thumbDIB.eraseDIB
            Set imgThumbnails(thumbIndex).thumbDIB = Nothing
        End If
        
        If Not (imgThumbnails(thumbIndex).thumbShadow Is Nothing) Then
            imgThumbnails(thumbIndex).thumbShadow.eraseDIB
            Set imgThumbnails(thumbIndex).thumbShadow = Nothing
        End If
        
        For i = thumbIndex To numOfThumbnails - 1
            Set imgThumbnails(i).thumbDIB = imgThumbnails(i + 1).thumbDIB
            Set imgThumbnails(i).thumbShadow = imgThumbnails(i + 1).thumbShadow
            imgThumbnails(i).indexInPDImages = imgThumbnails(i + 1).indexInPDImages
        Next i
        
        'Decrease the array size to erase the unneeded trailing entry
        numOfThumbnails = numOfThumbnails - 1
    
        If numOfThumbnails < 0 Then
            numOfThumbnails = 0
            curThumb = 0
        End If
        
        ReDim Preserve imgThumbnails(0 To numOfThumbnails) As thumbEntry
        
    End If
    
    'Because inactive images can be unloaded via the Win 7 taskbar, it is possible for our curThumb tracker to get out of sync.
    ' Update it now.
    For i = 0 To numOfThumbnails
        If imgThumbnails(i).indexInPDImages = g_CurrentImage Then
            curThumb = i
            Exit For
        Else
            curThumb = 0
        End If
    Next i
    
    'Redraw the toolbar to reflect these changes
    If refreshToolbar Then redrawToolbar

End Sub

'Given mouse coordinates over the form, return the thumbnail at that location.  If the cursor is not over a thumbnail,
' the function will return -1
Private Function getThumbAtPosition(ByVal x As Long, ByVal y As Long) As Long
    
    Dim thumbOffset As Long
    thumbOffset = hsThumbnails.Value
    
    If verticalLayout Then
        getThumbAtPosition = (y + thumbOffset) \ thumbHeight
        If getThumbAtPosition > (numOfThumbnails - 1) Then getThumbAtPosition = -1
    Else
        getThumbAtPosition = (x + thumbOffset) \ thumbWidth
        If getThumbAtPosition > (numOfThumbnails - 1) Then getThumbAtPosition = -1
    End If
    
End Function

'Given mouse coordinates over the form, specify whether the cursor is over the "close" icon area.
' RETURNS: thumbnail index if over a close icon on a thumbnail, -1 if not over a close icon.
Private Function getThumbWithCloseIconAtPosition(ByVal x As Long, ByVal y As Long) As Long
    Dim thumbnailNumber As Long
    Dim thumbnailStartOffsetX As Long
    Dim thumbnailStartOffsetY As Long
    Dim closeButtonStartOffsetX As Long
    Dim closeButtonStartOffsetY As Long
    Dim clickboundaryX As Long
    Dim clickBoundaryY As Long
    
    Dim thumbScrollOffset As Long
    thumbScrollOffset = hsThumbnails.Value
    
    getThumbWithCloseIconAtPosition = -1
    thumbnailNumber = getThumbAtPosition(x, y)
    
    If thumbnailNumber <> -1 Then
        If verticalLayout Then
            thumbnailStartOffsetX = 0
            thumbnailStartOffsetY = thumbHeight * thumbnailNumber - thumbScrollOffset
        Else
            thumbnailStartOffsetX = thumbWidth * thumbnailNumber - thumbScrollOffset
            thumbnailStartOffsetY = 0
        End If
        
        closeButtonStartOffsetX = thumbnailStartOffsetX + (thumbWidth - (FixDPI(thumbBorder) + m_CloseIconGray.getDIBWidth + FixDPI(2)))
        closeButtonStartOffsetY = thumbnailStartOffsetY + FixDPI(thumbBorder) + FixDPI(2)
        clickboundaryX = x - closeButtonStartOffsetX
        clickBoundaryY = y - closeButtonStartOffsetY
        
        If clickboundaryX >= 0 And clickboundaryX <= m_CloseIconGray.getDIBWidth Then
            If clickBoundaryY >= 0 And clickBoundaryY <= m_CloseIconGray.getDIBHeight Then
                getThumbWithCloseIconAtPosition = thumbnailNumber
            End If
        End If
    End If
    
End Function

'Given an x/y mouse coordinate, return TRUE if the coordinate falls over the form resize area.  Tabstrip alignment is automatically handled.
Private Function isMouseOverResizeBorder(ByVal mouseX As Single, ByVal mouseY As Single) As Boolean

    'How close does the mouse have to be to the form border to allow resizing?  We currently use 7 pixels, while accounting
    ' for DPI variance (e.g. 7 pixels at 96 dpi)
    Dim resizeBorderAllowance As Long
    resizeBorderAllowance = FixDPI(7)
    
    Select Case g_WindowManager.GetImageTabstripAlignment
    
        Case vbAlignLeft
            If (mouseY > 0) And (mouseY < Me.ScaleHeight) And (mouseX > Me.ScaleWidth - resizeBorderAllowance) Then isMouseOverResizeBorder = True
            
        Case vbAlignTop
            If (mouseX > 0) And (mouseX < Me.ScaleWidth) And (mouseY > Me.ScaleHeight - resizeBorderAllowance) Then isMouseOverResizeBorder = True
            
        Case vbAlignRight
            If (mouseY > 0) And (mouseY < Me.ScaleHeight) And (mouseX < resizeBorderAllowance) Then isMouseOverResizeBorder = True
            
        Case vbAlignBottom
            If (mouseX > 0) And (mouseX < Me.ScaleWidth) And (mouseY < resizeBorderAllowance) Then isMouseOverResizeBorder = True
            
    End Select

End Function

'Click events are automatically sorted by the cMouseEvents class.  It will also raise a _MouseUp event, but a parameter in that event notifies
' that _Click is also being raised; that allows us to specifically handle click-only behavior here (such as raising a context menu).
Private Sub cMouseEvents_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'Separate handling by button
    Select Case Button
    
        Case pdLeftButton
            
            'If the mouse is not over a close icon, attempt to activate that image now
            If m_CloseTriggeredOnThumbnail = -1 Then
            
                Dim potentialNewThumb As Long
                potentialNewThumb = getThumbAtPosition(x, y)
                
                'Notify the program that a new image has been selected; it will then bring that image to the foreground,
                ' which will automatically trigger a toolbar redraw.  Also, do not select the image if the user has been
                ' scrolling the list.
                If (potentialNewThumb >= 0) And (Not m_ScrollingOccured) Then
                    curThumb = potentialNewThumb
                    ActivatePDImage imgThumbnails(curThumb).indexInPDImages, "user clicked image thumbnail"
                End If
                
            'If the mouse *IS* over a close icon, close the image in question
            Else
               
                'm_CloseTriggeredOnThumbnail is set at MouseDown, if the mouse is over a close icon
                If getThumbWithCloseIconAtPosition(x, y) = m_CloseTriggeredOnThumbnail Then
                    
                    'fullPDImageUnload will take care of refreshing the UI, activating the next thumbnail if the active one is
                    ' closed, showing a dialog before closing an unsaved image, etc.
                    Image_Canvas_Handler.FullPDImageUnload imgThumbnails(m_CloseTriggeredOnThumbnail).indexInPDImages
                    
                End If
    
                'Reset the close identifier
                m_CloseTriggeredOnThumbnail = -1
                
            End If
            
        'Right button raises a context menu (potentially)
        Case pdRightButton
        
            ' If a thumbnail was right-clicked at mousedown, and mouseup happens on
            '   the same thumbnail, activate the image and show the context menu
            If m_RightClickedThumbnail <> -1 Then
                If m_RightClickedThumbnail = getThumbAtPosition(x, y) Then
                
                    'Activate the image, which triggers a redraw and resets all of PD's internal image tracking data
                    curThumb = m_RightClickedThumbnail
                    ActivatePDImage imgThumbnails(curThumb).indexInPDImages, "user right-clicked image thumbnail"
                     
                    'Enable various pop-up menu entries.  Wherever possible, we simply want to mimic the official PD menu, which saves
                    ' us having to supply our own heuristics for menu enablement.
                    mnuTabstripPopup(POP_SAVE).Enabled = FormMain.MnuFile(8).Enabled
                    mnuTabstripPopup(POP_SAVE_COPY).Enabled = FormMain.MnuFile(9).Enabled
                    mnuTabstripPopup(POP_SAVE_AS).Enabled = FormMain.MnuFile(10).Enabled
                    mnuTabstripPopup(POP_REVERT).Enabled = FormMain.MnuFile(11).Enabled
                    mnuTabstripPopup(POP_CLOSE).Enabled = FormMain.MnuFile(5).Enabled
                    
                    'Two special commands only appear in this menu: Open in Explorer, and Close Other Images
                    ' Use our own enablement heuristics for these.
                    
                    'Open in Explorer only works if the image is currently on-disk
                    mnuTabstripPopup(POP_OPEN_IN_EXPLORER).Enabled = (Len(pdImages(imgThumbnails(curThumb).indexInPDImages).locationOnDisk) > 0)
                    
                    'Close Other Images only works if more than one image is open.  We can determine this using the Next/Previous Image items
                    ' in the Window menu
                    mnuTabstripPopup(POP_CLOSE).Enabled = FormMain.MnuWindow(5).Enabled
                    
                    'Raise the context menu
                    Me.PopupMenu mnuImageTabsContext, x:=x, y:=y
                    
                    'Reset the tabstrip, then exit
                    m_RightClickedThumbnail = -1
                    forceRedraw
                    Exit Sub
                    
                End If
            End If
        
    End Select
    
End Sub

'When the left mouse button is pressed, activate click-to-drag mode for scrolling the tabstrip window
Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'On left-button presses, make a note of the initial mouse position
    If Button = vbLeftButton Then
    
        m_MouseDown = True
        m_InitX = x
        m_InitY = y
        m_MouseDistanceTraveled = 0
        m_InitOffset = hsThumbnails.Value
        
        'Detect close icon click, and store the clicked thumbnail
        m_CloseTriggeredOnThumbnail = getThumbWithCloseIconAtPosition(x, y)
        
        'We must also detect if the mouse is over the edge of the form that allows live-resizing.  (This varies by tabstrip orientation, obviously.)
        m_MouseInResizeTerritory = isMouseOverResizeBorder(x, y)
        
    ElseIf Button = vbRightButton Then
        m_RightClickedThumbnail = getThumbAtPosition(x, y)
    End If
    
    'Reset the "resize in progress" tracker
    weAreResponsibleForResize = False
    
    'Reset the "scrolling occured" tracker
    m_ScrollingOccured = False

End Sub

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    g_MouseOverImageTabstrip = True
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    g_MouseOverImageTabstrip = False
    
    If curThumbHover <> -1 Then
        curThumbHover = -1
        redrawToolbar
    End If
    
    cMouseEvents.setSystemCursor IDC_ARROW

End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Note that the mouse is currently over the tabstrip
    g_MouseOverImageTabstrip = True
    
    'We require a few mouse movements to fire before doing anything; otherwise this function will fire constantly.
    m_MouseDistanceTraveled = m_MouseDistanceTraveled + 1
    
    'We handle several different _MouseMove scenarios, in this order:
    ' 1) If the mouse is near the resizable edge of the form, and the left button is depressed, activate live resizing.
    ' 2) If a button is depressed, activate tabstrip scrolling (if the list is long enough)
    ' 3) If no buttons are depressed, hover the image at the current position (if any)
        
    'Check mouse button state; if it's down, check for resize or scrolling of the image list
    If m_MouseDown Then
        
        If m_MouseInResizeTerritory Then
                
            If (Button = vbLeftButton) Then
            
                'Figure out which resize message to send to Windows; this varies by tabstrip orientation (obviously)
                Dim hitCode As Long
    
                Select Case g_WindowManager.GetImageTabstripAlignment
                
                    Case vbAlignLeft
                        hitCode = HTRIGHT
                    
                    Case vbAlignTop
                        hitCode = HTBOTTOM
                    
                    Case vbAlignRight
                        hitCode = HTLEFT
                    
                    Case vbAlignBottom
                        hitCode = HTTOP
                
                End Select
                
                'Initiate resizing, and set a form-level marker so that other functions know we're responsible for any resize-related events
                weAreResponsibleForResize = True
                ReleaseCapture
                SendMessage Me.hWnd, WM_NCLBUTTONDOWN, hitCode, ByVal 0&
                
            End If
        
        'The mouse is not in resize territory.  This means the user is click-dragging to scroll a long list.
        Else
            
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
    
        'We want to highlight a close icon, if it's being hovered
        m_CloseIconHovered = getThumbWithCloseIconAtPosition(x, y)
        
        Dim oldThumbHover As Long
        oldThumbHover = curThumbHover
        
        'Retrieve the thumbnail at this position, and change the mouse pointer accordingly
        curThumbHover = getThumbAtPosition(x, y)
                
        'To prevent flickering, only update the tooltip when absolutely necessary
        If curThumbHover <> oldThumbHover Then
        
            'If the cursor is over a thumbnail, update the tooltip to display that image's filename
            If curThumbHover <> -1 Then
                        
                If Len(pdImages(imgThumbnails(curThumbHover).indexInPDImages).locationOnDisk) <> 0 Then
                    toolTipManager.setTooltip Me.hWnd, Me.hWnd, pdImages(imgThumbnails(curThumbHover).indexInPDImages).locationOnDisk, pdImages(imgThumbnails(curThumbHover).indexInPDImages).originalFileNameAndExtension
                Else
                    toolTipManager.setTooltip Me.hWnd, Me.hWnd, "Once this image has been saved to disk, its filename will appear here.", "This image does not have a filename."
                End If
            
            'The cursor is not over a thumbnail; let the user know they can hover if they want more information.
            Else
            
                toolTipManager.setTooltip Me.hWnd, Me.hWnd, "Hover an image thumbnail to see its name and current file location.", ""
            
            End If
            
        End If
        
    End If
    
    'Set a mouse pointer according to the handling above
    If isMouseOverResizeBorder(x, y) Then
    
        If verticalLayout Then
            cMouseEvents.setSystemCursor IDC_SIZEWE
        Else
            cMouseEvents.setSystemCursor IDC_SIZENS
        End If
        
    Else
    
        'Display a hand cursor if over an image; default cursor otherwise
        If curThumbHover = -1 Then cMouseEvents.setSystemCursor IDC_ARROW Else cMouseEvents.setSystemCursor IDC_HAND
    
    End If
    
    'Regardless of what happened above, redraw the toolbar to reflect any changes
    redrawToolbar
    
End Sub

Private Sub cMouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    
    'cMouseEvents is nice enough to auto-detect clicks and raise the _Click event accordingly.  This saves us having to deal with certain
    ' potential outcomes here.
        
    'Release mouse tracking, if any
    If m_MouseDown Then
        m_MouseDown = False
        m_InitX = 0
        m_InitY = 0
        m_MouseDistanceTraveled = 0
    End If

End Sub

Public Sub cMouseEvents_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

    'Horizontal scrolling - only trigger it if the horizontal scroll bar is actually visible
    If m_ListScrollable Then
  
        If scrollAmount > 0 Then
            
            If hsThumbnails.Value + hsThumbnails.LargeChange > hsThumbnails.Max Then
                hsThumbnails.Value = hsThumbnails.Max
            Else
                hsThumbnails.Value = hsThumbnails.Value + hsThumbnails.LargeChange
            End If
            
            curThumbHover = getThumbAtPosition(x, y)
            redrawToolbar
        
        ElseIf scrollAmount < 0 Then
            
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

Public Sub cMouseEvents_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

    'Vertical scrolling - only trigger it if the horizontal scroll bar is actually visible
    If m_ListScrollable Then
  
        If scrollAmount < 0 Then
            
            If hsThumbnails.Value + hsThumbnails.LargeChange > hsThumbnails.Max Then
                hsThumbnails.Value = hsThumbnails.Max
            Else
                hsThumbnails.Value = hsThumbnails.Value + hsThumbnails.LargeChange
            End If
            
            curThumbHover = getThumbAtPosition(x, y)
            redrawToolbar
        
        ElseIf scrollAmount > 0 Then
            
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

    'Initialize the back buffer
    Set bufferDIB = New pdDIB

    'Reset the thumbnail array
    numOfThumbnails = 0
    ReDim imgThumbnails(0 To numOfThumbnails) As thumbEntry
    
    'Enable mousewheel scrolling
    Set cMouseEvents = New pdInputMouse
    cMouseEvents.addInputTracker Me.hWnd, True, True, , True
    cMouseEvents.setSystemCursor IDC_HAND
    
    'Detect initial alignment
    If (g_WindowManager.GetImageTabstripAlignment = vbAlignLeft) Or (g_WindowManager.GetImageTabstripAlignment = vbAlignRight) Then
        verticalLayout = True
    Else
        verticalLayout = False
    End If
    
    'Set default thumbnail sizes
    If verticalLayout Then
        thumbWidth = g_WindowManager.GetClientWidth(Me.hWnd)
        thumbHeight = thumbWidth
    Else
        thumbHeight = g_WindowManager.GetClientHeight(Me.hWnd)
        thumbWidth = thumbHeight
    End If
    
    'Compensate for the presence of the 2px border along the edge of the tabstrip
    thumbWidth = thumbWidth - 2
    thumbHeight = thumbHeight - 2
    
    'Retrieve the unsaved image notification icon from the resource file
    Set unsavedChangesDIB = New pdDIB
    loadResourceToDIB "NTFY_UNSAVED", unsavedChangesDIB
    
    'Retrieve all PNGs necessary to render the "close by hovering" X that appears
    Set m_CloseIconRed = New pdDIB
    loadResourceToDIB "CLOSE_MINI_RED", m_CloseIconRed
    
    Set m_CloseIconGray = New pdDIB
    loadResourceToDIB "CLOSE_MINI_GRAY", m_CloseIconGray
    
    'Update the drop-shadow blur radius to account for DPI
    shadowBlurRadius = FixDPI(2)
    
    'Generate a drop-shadow for the X.  (We can use the same one for both red and gray, obviously.)
    Set m_CloseIconShadow = New pdDIB
    Filters_Layers.createShadowDIB m_CloseIconGray, m_CloseIconShadow
    m_CloseIconShadow.setAlphaPremultiplication False
    
    'Pad and blur the drop-shadow
    Dim tmpLUT() As Byte
    
    Dim cFilter As pdFilterLUT
    Set cFilter = New pdFilterLUT
    cFilter.fillLUT_Invert tmpLUT
    
    padDIB m_CloseIconShadow, FixDPI(thumbBorder)
    quickBlurDIB m_CloseIconShadow, FixDPI(2), False
    cFilter.applyLUTToAllColorChannels m_CloseIconShadow, tmpLUT, True
    
    m_CloseIconShadow.setAlphaPremultiplication True
    
    ' Track the last thumbnail whose close icon has been clicked.
    ' -1 means no close icon has been clicked yet
    m_CloseTriggeredOnThumbnail = -1
    
    ' Track the last right-clicked thumbnail.
    m_RightClickedThumbnail = -1
        
    'If the tabstrip ever becomes long enough to scroll, this will be set to TRUE
    m_ListScrollable = False
    
    'Activate the custom tooltip handler
    Set toolTipManager = New pdToolTip
    
    'As a final step, redraw everything against the current theme.
    UpdateAgainstCurrentTheme
    
End Sub

'(This code is copied from FormMain's OLEDragDrop event - please mirror any changes there)
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub
    
    'Use the external function (in the clipboard handler, as the code is roughly identical to clipboard pasting)
    ' to load the OLE source.
    Clipboard_Handler.loadImageFromDragDrop Data, Effect, False

End Sub

'(This code is copied from FormMain's OLEDragOver event - please mirror any changes there)
Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

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

'Any time this window is resized, we need to recreate the thumbnail display
Private Sub Form_Resize()

    'Detect alignment changes (if any)
    If (g_WindowManager.GetImageTabstripAlignment = vbAlignLeft) Or (g_WindowManager.GetImageTabstripAlignment = vbAlignRight) Then
        verticalLayout = True
    Else
        verticalLayout = False
    End If

    Dim i As Long

    'If the tabstrip is horizontal and the window's height is changing, we need to recreate all image thumbnails
    If ((Not verticalLayout) And (thumbHeight <> g_WindowManager.GetClientHeight(Me.hWnd) - 2)) Then
        
        thumbHeight = g_WindowManager.GetClientHeight(Me.hWnd) - 2
    
        For i = 0 To numOfThumbnails - 1
            imgThumbnails(i).thumbDIB.eraseDIB
            pdImages(imgThumbnails(i).indexInPDImages).requestThumbnail imgThumbnails(i).thumbDIB, thumbHeight - (FixDPI(thumbBorder) * 2)
            If g_InterfacePerformance <> PD_PERF_FASTEST Then updateShadowDIB i
        Next i
    
    End If
    
    'If the tabstrip is vertical and the window's with is changing, we need to recreate all image thumbnails
    If (verticalLayout And (thumbWidth <> g_WindowManager.GetClientWidth(Me.hWnd) - 2)) Then
    
        thumbWidth = g_WindowManager.GetClientWidth(Me.hWnd) - 2
        
        For i = 0 To numOfThumbnails - 1
            imgThumbnails(i).thumbDIB.eraseDIB
            pdImages(imgThumbnails(i).indexInPDImages).requestThumbnail imgThumbnails(i).thumbDIB, thumbWidth - (FixDPI(thumbBorder) * 2)
            If g_InterfacePerformance <> PD_PERF_FASTEST Then updateShadowDIB i
        Next i
    
    End If
    
    'Update thumbnail sizes
    If verticalLayout Then
        thumbWidth = g_WindowManager.GetClientWidth(Me.hWnd) - 2
        thumbHeight = thumbWidth
    Else
        thumbHeight = g_WindowManager.GetClientHeight(Me.hWnd) - 2
        thumbWidth = thumbHeight
    End If
        
    'Create a background buffer the same size as this window
    m_BufferWidth = g_WindowManager.GetClientWidth(Me.hWnd)
    m_BufferHeight = g_WindowManager.GetClientHeight(Me.hWnd)
    
    'Redraw the toolbar
    redrawToolbar
    
    'Notify the window manager that the tab strip has been resized; it will resize image windows to match
    'If Not weAreResponsibleForResize Then
    g_WindowManager.NotifyImageTabStripResized
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    g_WindowManager.UnregisterForm Me
    Set cMouseEvents = Nothing
End Sub

'Whenever a function wants to render the current toolbar, it may do so here
Private Sub redrawToolbar(Optional ByVal fitCurrentThumbOnScreen As Boolean = False)

    'Recreate the toolbar buffer
    bufferDIB.createBlank m_BufferWidth, m_BufferHeight, 24, ConvertSystemColor(vb3DShadow)
    
    If numOfThumbnails > 0 Then
        
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
        maxThumbSize = constrainingDimension * numOfThumbnails - 1
        
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
            
            'If requested, fit the currently active thumbnail on-screen
            If fitCurrentThumbOnScreen Then fitThumbnailOnscreen curThumb
            
        End If
        
        'Determine a scrollbar offset as necessary
        Dim scrollOffset As Long
        scrollOffset = hsThumbnails.Value
        
        'Render each thumbnail block
        Dim i As Long
        For i = 0 To numOfThumbnails - 1
            If verticalLayout Then
                If g_WindowManager.GetImageTabstripAlignment = vbAlignLeft Then
                    renderThumbTab i, 0, (i * thumbHeight) - scrollOffset
                Else
                    renderThumbTab i, 2, (i * thumbHeight) - scrollOffset
                End If
            Else
                If g_WindowManager.GetImageTabstripAlignment = vbAlignTop Then
                    renderThumbTab i, (i * thumbWidth) - scrollOffset, 0
                Else
                    renderThumbTab i, (i * thumbWidth) - scrollOffset, 2
                End If
            End If
        Next i
        
        'Eventually we'll do something nicer, but for now, draw a line across the edge of the tabstrip nearest the image.
        Select Case g_WindowManager.GetImageTabstripAlignment
        
            Case vbAlignLeft
                GDIPlusDrawLineToDC bufferDIB.getDIBDC, m_BufferWidth - 1, 0, m_BufferWidth - 1, m_BufferHeight, ConvertSystemColor(vb3DLight), 255, 2, False
            
            Case vbAlignTop
                GDIPlusDrawLineToDC bufferDIB.getDIBDC, 0, m_BufferHeight - 1, m_BufferWidth, m_BufferHeight - 1, ConvertSystemColor(vb3DLight), 255, 2, False
            
            Case vbAlignRight
                GDIPlusDrawLineToDC bufferDIB.getDIBDC, 1, 0, 1, m_BufferHeight, ConvertSystemColor(vb3DLight), 255, 2, False
            
            Case vbAlignBottom
                GDIPlusDrawLineToDC bufferDIB.getDIBDC, 0, 1, m_BufferWidth, 1, ConvertSystemColor(vb3DLight), 255, 2, False
        
        End Select
        
    End If
    
    'Activate color management for our form
    AssignDefaultColorProfileToObject Me.hWnd, Me.hDC
    TurnOnColorManagementForDC Me.hDC
    
    'Copy the buffer to the form
    BitBlt Me.hDC, 0, 0, m_BufferWidth, m_BufferHeight, bufferDIB.getDIBDC, 0, 0, vbSrcCopy
    Me.Picture = Me.Image
    Me.Refresh
    
End Sub
    
'Render a given thumbnail onto the background form at the specified offset
Private Sub renderThumbTab(ByVal thumbIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current tab if it will be visible
    Dim tabVisible As Boolean
    tabVisible = False
    
    If verticalLayout Then
        If ((offsetY + thumbHeight) > 0) And (offsetY < m_BufferHeight) Then tabVisible = True
    Else
        If ((offsetX + thumbWidth) > 0) And (offsetX < m_BufferWidth) Then tabVisible = True
    End If
    
    If tabVisible Then
    
        Dim tmpRect As RECTL
        Dim hBrush As Long
    
        'If this thumbnail has been selected, draw the background with the system's current selection color
        If (thumbIndex = curThumb) Then
            SetRect tmpRect, offsetX, offsetY, offsetX + thumbWidth, offsetY + thumbHeight
            hBrush = CreateSolidBrush(ConvertSystemColor(vb3DLight))
            FillRect bufferDIB.getDIBDC, tmpRect, hBrush
            DeleteObject hBrush
        End If
        
        'If the current thumbnail is highlighted but not selected, simply render the border with a highlight
        If (thumbIndex <> curThumb) And (thumbIndex = curThumbHover) Then
            SetRect tmpRect, offsetX, offsetY, offsetX + thumbWidth, offsetY + thumbHeight
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FrameRect bufferDIB.getDIBDC, tmpRect, hBrush
            SetRect tmpRect, tmpRect.Left + 1, tmpRect.Top + 1, tmpRect.Right - 1, tmpRect.Bottom - 1
            FrameRect bufferDIB.getDIBDC, tmpRect, hBrush
            DeleteObject hBrush
        End If
    
        'Render the matching thumbnail shadow and thumbnail into this block
        If g_InterfacePerformance <> PD_PERF_FASTEST Then imgThumbnails(thumbIndex).thumbShadow.alphaBlendToDC bufferDIB.getDIBDC, 192, offsetX, offsetY + FixDPI(1)
        imgThumbnails(thumbIndex).thumbDIB.alphaBlendToDC bufferDIB.getDIBDC, 255, offsetX + FixDPI(thumbBorder), offsetY + FixDPI(thumbBorder)
        
        'If the parent image has unsaved changes, also render a notification icon.
        ' (When the program is shutting down, this value may not be available, so we must check it first.
        If Not (pdImages(imgThumbnails(thumbIndex).indexInPDImages) Is Nothing) Then
            If Not pdImages(imgThumbnails(thumbIndex).indexInPDImages).getSaveState(pdSE_AnySave) Then
                unsavedChangesDIB.alphaBlendToDC bufferDIB.getDIBDC, 230, offsetX + FixDPI(thumbBorder) + FixDPI(2), offsetY + thumbHeight - FixDPI(thumbBorder) - unsavedChangesDIB.getDIBHeight - FixDPI(2)
            End If
        End If
        
        'If this image is being hovered over, show the close icon.  (Drop shadow gets rendered first, at a slight offset.)
        If (thumbIndex = curThumbHover) Then
        
            'Render a shadow, regardless
            m_CloseIconShadow.alphaBlendToDC bufferDIB.getDIBDC, 230, offsetX + (thumbWidth - (FixDPI(thumbBorder) * 2 + m_CloseIconRed.getDIBWidth + FixDPI(2))), offsetY + FixDPI(2)
            
            'Render a red icon if the close icon is actually hovered; gray, otherwise
            If (thumbIndex = m_CloseIconHovered) Then
                m_CloseIconRed.alphaBlendToDC bufferDIB.getDIBDC, 230, offsetX + (thumbWidth - (FixDPI(thumbBorder) + m_CloseIconRed.getDIBWidth + FixDPI(2))), offsetY + FixDPI(thumbBorder) + FixDPI(2)
            Else
                m_CloseIconGray.alphaBlendToDC bufferDIB.getDIBDC, 230, offsetX + (thumbWidth - (FixDPI(thumbBorder) + m_CloseIconRed.getDIBWidth + FixDPI(2))), offsetY + FixDPI(thumbBorder) + FixDPI(2)
            End If
            
        End If
        
    End If

End Sub

'Whenever a thumbnail has been updated, this sub must be called to regenerate its drop-shadow
Private Sub updateShadowDIB(ByVal imgThumbnailIndex As Long)
    
    'Create a shadow DIB
    imgThumbnails(imgThumbnailIndex).thumbShadow.eraseDIB
    createShadowDIB imgThumbnails(imgThumbnailIndex).thumbDIB, imgThumbnails(imgThumbnailIndex).thumbShadow
    
    'Pad and blur the DIB
    padDIB imgThumbnails(imgThumbnailIndex).thumbShadow, FixDPI(thumbBorder)
    quickBlurDIB imgThumbnails(imgThumbnailIndex).thumbShadow, shadowBlurRadius
    
    'Apply premultiplied alpha (so we can more quickly AlphaBlend the resulting image to the tabstrip)
    imgThumbnails(imgThumbnailIndex).thumbShadow.setAlphaPremultiplication True
    
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
    MakeFormPretty Me
End Sub

'All popup menu clicks are handled here
Private Sub mnuTabstripPopup_Click(Index As Integer)

    Select Case Index
        
        'Save
        Case 0
            File_Menu.MenuSave imgThumbnails(m_RightClickedThumbnail).indexInPDImages
        
        'Save copy (lossless)
        Case 1
            File_Menu.MenuSaveLosslessCopy imgThumbnails(m_RightClickedThumbnail).indexInPDImages
        
        'Save as
        Case 2
            File_Menu.MenuSaveAs imgThumbnails(m_RightClickedThumbnail).indexInPDImages
        
        'Revert
        Case 3
            Dim imageToRevert As Long
            imageToRevert = imgThumbnails(m_RightClickedThumbnail).indexInPDImages
            
            pdImages(imageToRevert).undoManager.revertToLastSavedState
                        
            'Also, redraw the current child form icon
            createCustomFormIcon pdImages(imageToRevert)
            notifyUpdatedImage imageToRevert
        
        '(separator)
        Case 4
        
        'Open location in Explorer
        Case 5
            Dim filePath As String, shellCommand As String
            filePath = pdImages(imgThumbnails(m_RightClickedThumbnail).indexInPDImages).locationOnDisk
            shellCommand = "explorer.exe /select,""" & filePath & """"
            Shell shellCommand, vbNormalFocus
        
        '(separator)
        Case 6
        
        'Close
        Case 7
            Image_Canvas_Handler.FullPDImageUnload imgThumbnails(m_RightClickedThumbnail).indexInPDImages
        
        'Close all but this
        Case 8
            Dim lastImageIndex As Long
            Dim rightclickedImageIndex As Long
            Dim i As Long
            
            lastImageIndex = UBound(pdImages)
            rightclickedImageIndex = imgThumbnails(m_RightClickedThumbnail).indexInPDImages
            
            For i = 0 To lastImageIndex
                If i <> rightclickedImageIndex And (Not pdImages(i) Is Nothing) Then
                    FullPDImageUnload i
                End If
            Next i
    
    End Select

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
    
    'Redraw the tabstrip
    redrawToolbar
    
End Sub
