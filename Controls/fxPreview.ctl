VERSION 5.00
Begin VB.UserControl fxPreviewCtl 
   AccessKeys      =   "T"
   BackColor       =   &H80000005&
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
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
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
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
      MouseIcon       =   "fxPreview.ctx":0312
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
'Copyright ©2012-2013 by Tanner Helland
'Created: 10/January/13
'Last updated: 26/July/13
'Last update: use Alt+T as an accelerator to toggle between original and preview image
'
'For the first decade of its life, PhotoDemon relied on simple picture boxes for rendering its effect previews.
' This worked well enough when there were only a handful of tools available, but as the complexity of the program
' - and its various effects and tools - has grown, it has become more and more painful to update the preview
' system, because any changes have to be mirrored across a huge number of forms.
'
'Thus, this control was born.  It is now used on every single effect form in place of a regular picture box.  This
' allows me to add preview-related features just once - to the base control - and every tool will automatically
' reap the benefits.
'
'At present, there isn't much to the control.  It is capable of storing a copy of the original image and any
' filter-modified versions of the image.  The user can toggle between these by using the command link below the
' main picture box, or by pressing Alt+T.  This replaces the side-by-side "before and after" of past versions.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Has this control been given a copy of the original image?
Dim m_HasOriginal As Boolean, m_HasFX As Boolean

Dim originalImage As pdLayer, fxImage As pdLayer

'The control's current state: whether it is showing the original image or the fx preview
Dim curImageState As Boolean

'Use this to supply the preview with a copy of the original image's data.  The preview object can use this to display
' the original image when the user clicks the "show original image" link.
Public Sub setOriginalImage(ByRef srcLayer As pdLayer)

    'Note that we have a copy of the original image, so the calling function doesn't attempt to supply it again
    m_HasOriginal = True
    
    'Make a copy of the layer passed in
    If (originalImage Is Nothing) Then Set originalImage = New pdLayer
    
    originalImage.eraseLayer
    originalImage.createFromExistingLayer srcLayer

End Sub

'Use this to supply the object with a copy of the processed image's data.  The preview object can use this to display
' the processed image again if the user clicks the "show original image" link, then clicks it again.
Public Sub setFXImage(ByRef srcLayer As pdLayer)

    'Note that we have a copy of the original image, so the calling function doesn't attempt to supply it again
    m_HasFX = True
    
    'Make a copy of the layer passed in
    If (fxImage Is Nothing) Then Set fxImage = New pdLayer
    
    fxImage.eraseLayer
    fxImage.createFromExistingLayer srcLayer
    
    'If the user was previously examining the original image, reset the label caption to match the new preview
    If Not curImageState Then
        lblBeforeToggle.Caption = g_Language.TranslateMessage("show original image") & " (alt+t) "
        curImageState = True
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

'When the control's access key is pressed (alt+t) , toggle the original/current image
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    lblBeforeToggle_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    
    'Keep the control's backcolor in sync with the parent object
    If UCase$(PropertyName) = "BACKCOLOR" Then
        backColor = Ambient.backColor
    End If

End Sub

Private Sub UserControl_Initialize()
    
    'Give the user control the same font as the rest of the program
    lblBeforeToggle.FontName = g_InterfaceFont
    
    curImageState = True
    
    setArrowCursorToHwnd UserControl.hWnd
    setArrowCursorToHwnd picPreview.hWnd
            
End Sub

'Initialize our effect preview control
Private Sub UserControl_InitProperties()
    
    'Set the background of the fxPreview to match the background of our parent object
    backColor = Ambient.backColor
    
    'Mark the original image as having NOT been set
    m_HasOriginal = False
    
End Sub

'Redraw the user control after it has been resized
Private Sub UserControl_Resize()
    redrawControl
End Sub

Private Sub UserControl_Show()
    'Translate the user control text
    If Ambient.UserMode Then
        lblBeforeToggle.Caption = g_Language.TranslateMessage("show original image") & " (alt+t) "
    Else
        lblBeforeToggle.Caption = "show original image (alt+t) "
    End If
End Sub

Private Sub UserControl_Terminate()

    'Release any image objects that may have been created
    If Not (originalImage Is Nothing) Then originalImage.eraseLayer
    If Not (fxImage Is Nothing) Then fxImage.eraseLayer
    
End Sub

'After a resize or paint request, update the layout of our control
Private Sub redrawControl()
    
    'Always make the preview picture box the width of the user control (at present)
    picPreview.Width = ScaleWidth
    
    'Adjust the preview picture box's height to be just above the "show original image" link
    lblBeforeToggle.Top = ScaleHeight - 24
    picPreview.Height = lblBeforeToggle.Top - (ScaleHeight - (lblBeforeToggle.Height + lblBeforeToggle.Top))
        
End Sub
