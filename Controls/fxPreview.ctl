VERSION 5.00
Begin VB.UserControl fxPreviewCtl 
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
      Height          =   240
      Left            =   120
      MouseIcon       =   "fxPreview.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5280
      Width           =   3045
   End
End
Attribute VB_Name = "fxPreviewCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Has this control been given a copy of the original image?
Dim m_HasOriginal As Boolean, m_HasFX As Boolean

Dim originalImage As pdLayer, fxImage As pdLayer

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
    If lblBeforeToggle.Caption <> "show original image" Then lblBeforeToggle.Caption = "show original image"

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
    If lblBeforeToggle.Caption = "show original image" Then lblBeforeToggle.Caption = "show effect preview" Else lblBeforeToggle.Caption = "show original image"
    lblBeforeToggle.Refresh
    
    'Update the image to match the new caption
    If lblBeforeToggle.Caption <> "show original image" Then
        If m_HasOriginal Then originalImage.renderToPictureBox picPreview
    Else
        
        If m_HasFX Then
            fxImage.renderToPictureBox picPreview
        Else
            If m_HasOriginal Then originalImage.renderToPictureBox picPreview
        End If
    End If
    
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    
    'Keep the control's backcolor in sync with the parent object
    If UCase$(PropertyName) = "BACKCOLOR" Then
        backColor = Ambient.backColor
    End If

End Sub

Private Sub UserControl_Initialize()
    
    'Give the user control the same font as the rest of the program
    If g_IsVistaOrLater And g_UseFancyFonts Then
        lblBeforeToggle.FontName = "Segoe UI"
    Else
        lblBeforeToggle.FontName = "Tahoma"
    End If
    
End Sub

'Initialize our effect preview control
Private Sub UserControl_InitProperties()
    
    'Set the background of the fxPreview to match the background of our parent object
    backColor = Ambient.backColor
    
    'Mark the original image as having NOT been set
    m_HasOriginal = False
    
End Sub

'Redraw the user control
Private Sub UserControl_Paint()
    'redrawControl
End Sub

Private Sub UserControl_Resize()
    redrawControl
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
    
    Dim newPreviewHeight As Long
    newPreviewHeight = lblBeforeToggle.Top - (ScaleHeight - (lblBeforeToggle.Height + lblBeforeToggle.Top))
    
    picPreview.Height = newPreviewHeight
    
End Sub
