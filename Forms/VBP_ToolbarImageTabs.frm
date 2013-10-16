VERSION 5.00
Begin VB.Form toolbar_ImageTabs 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13710
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
   ScaleHeight     =   68
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   914
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsThumbnails 
      Height          =   255
      Left            =   0
      Max             =   10
      TabIndex        =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   13695
   End
   Begin VB.Line lineSeparator 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   2000
      Y1              =   67
      Y2              =   67
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
'Last updated: 15/October/13
'Last update: initial implementation
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
    indexInPDImages As Long
End Type

Private imgThumbnails() As thumbEntry
Private numOfThumbnails As Long

Private Const THUMB_TAB_WIDTH As Long = 68
Private Const THUMB_TAB_HEIGHT As Long = 68

'The back buffer we use to hold the thumbnail display
Private bufferLayer As pdLayer
Private m_BufferWidth As Long, m_BufferHeight As Long

'An outside class provides access to mousewheel events for scrolling the filter view
Private WithEvents cMouseEvents As bluMouseEvents
Attribute cMouseEvents.VB_VarHelpID = -1

'The currently selected and currently hovered thumbnail
Private curThumb As Long, curThumbHover As Long

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
            pdImages(pdImagesIndex).requestThumbnail imgThumbnails(i).thumbLayer, 64
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
    pdImages(pdImagesIndex).requestThumbnail imgThumbnails(numOfThumbnails).thumbLayer, 64
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
    imgThumbnails(thumbIndex).thumbLayer.eraseLayer
    Set imgThumbnails(thumbIndex).thumbLayer = Nothing
    
    For i = thumbIndex To numOfThumbnails - 1
        Set imgThumbnails(i).thumbLayer = imgThumbnails(i + 1).thumbLayer
        imgThumbnails(i).indexInPDImages = imgThumbnails(i + 1).indexInPDImages
    Next i
    
    'Decrease the array size to erase the unneeded trailing entry
    numOfThumbnails = numOfThumbnails - 1
    ReDim Preserve imgThumbnails(0 To numOfThumbnails) As thumbEntry
 
    'Redraw the toolbar to reflect these changes
    redrawToolbar

End Sub

'Given mouse coordinates over the form, return the thumbnail at that location.  If the cursor is not over a thumbnail,
' the function will return -1
Private Function getThumbAtPosition(ByVal x As Long, ByVal y As Long) As Long
    
    Dim hOffset As Long
    hOffset = hsThumbnails.Value
    
    getThumbAtPosition = (x + hOffset) \ fixDPI(THUMB_TAB_WIDTH)
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
    curThumbHover = -1
    cMouseEvents.MousePointer = 0
End Sub

Private Sub Form_Load()

    'Reset the thumbnail array
    numOfThumbnails = 0
    ReDim imgThumbnails(0 To numOfThumbnails) As thumbEntry
    
    'Enable mousewheel scrolling
    Set cMouseEvents = New bluMouseEvents
    cMouseEvents.Attach Me.hWnd
    cMouseEvents.MousePointer = IDC_HAND
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    curThumb = getThumbAtPosition(x, y)
    
    'Notify the program that a new image has been selected; it will then bring that image to the foreground,
    ' which will automatically trigger a toolbar redraw
    pdImages(imgThumbnails(curThumb).indexInPDImages).containingForm.ActivateWorkaround
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    curThumbHover = getThumbAtPosition(x, y)
    If curThumbHover = -1 Then cMouseEvents.MousePointer = 0 Else cMouseEvents.MousePointer = IDC_HAND
    redrawToolbar
End Sub

'Any time this window is resized, we need to recreate the thumbnail display
Private Sub Form_Resize()

    'Create a background buffer the same size as this window
    m_BufferWidth = g_WindowManager.getClientWidth(Me.hWnd)
    m_BufferHeight = g_WindowManager.getClientHeight(Me.hWnd)
    
    'Resize the scroll bar
    hsThumbnails.Width = m_BufferWidth
    
    'Redraw the toolbar
    redrawToolbar
    
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
    bufferLayer.createBlank m_BufferWidth, m_BufferHeight, 24, g_CanvasBackground
    
    'Determine if the horizontal scrollbar needs to be visible or not
    Dim maxThumbSize As Long
    maxThumbSize = fixDPIFloat(THUMB_TAB_WIDTH) * numOfThumbnails - 1
    
    If maxThumbSize < m_BufferWidth Then
        hsThumbnails.Value = 0
        hsThumbnails.Visible = False
    Else
        hsThumbnails.Visible = True
        hsThumbnails.Max = maxThumbSize - m_BufferWidth
    End If
    
    'Determine a scrollbar offset as necessary
    Dim scrollOffset As Long
    scrollOffset = hsThumbnails.Value
    
    'Render each thumbnail block
    Dim i As Long
    For i = 0 To numOfThumbnails - 1
        renderThumbTab i, fixDPI(i * THUMB_TAB_WIDTH) - scrollOffset - fixDPI(2), 0
    Next i
    
    'Copy the buffer to the form
    BitBlt Me.hDC, 0, 0, m_BufferWidth, m_BufferHeight, bufferLayer.getLayerDC, 0, 0, vbSrcCopy
    Me.Picture = Me.Image
    If Me.Visible Then Me.Refresh
    
End Sub
    
'Render a given thumbnail onto the background form at the specified offset
Private Sub renderThumbTab(ByVal thumbIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current tab if it will be visible
    If ((offsetX + fixDPI(THUMB_TAB_WIDTH)) > 0) And (offsetX < m_BufferWidth) Then
    
        Dim tmpRect As RECT
        Dim hBrush As Long
    
        'If this thumbnail has been selected, draw the background with the system's current selection color
        If thumbIndex = curThumb Then
            SetRect tmpRect, offsetX, offsetY, offsetX + fixDPI(THUMB_TAB_WIDTH), offsetY + fixDPI(THUMB_TAB_HEIGHT)
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FillRect bufferLayer.getLayerDC, tmpRect, hBrush
            DeleteObject hBrush
        End If
        
        'If the current thumbnail is highlighted but not selected, simply render the border with a highlight
        If (thumbIndex <> curThumb) And (thumbIndex = curThumbHover) Then
            SetRect tmpRect, offsetX, offsetY, offsetX + fixDPI(THUMB_TAB_WIDTH), offsetY + fixDPI(THUMB_TAB_HEIGHT)
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FrameRect bufferLayer.getLayerDC, tmpRect, hBrush
            DeleteObject hBrush
        End If
    
        'Render the matching thumbnail into this block
        imgThumbnails(thumbIndex).thumbLayer.alphaBlendToDC bufferLayer.getLayerDC, 255, offsetX + fixDPI(1), offsetY + fixDPI(1)
        
    End If

End Sub

Private Sub hsThumbnails_Change()
    redrawToolbar
End Sub

Private Sub hsThumbnails_Scroll()
    redrawToolbar
End Sub
