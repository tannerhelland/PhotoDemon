VERSION 5.00
Begin VB.Form FormTile 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tile Image"
   ClientHeight    =   2310
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtWidth 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Text            =   "N/A"
      Top             =   1020
      Width           =   855
   End
   Begin VB.TextBox TxtHeight 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   4080
      TabIndex        =   6
      Text            =   "N/A"
      Top             =   1020
      Width           =   855
   End
   Begin VB.VScrollBar VSWidth 
      Height          =   420
      Left            =   2070
      Max             =   32766
      Min             =   1
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Value           =   15000
      Width           =   270
   End
   Begin VB.VScrollBar VSHeight 
      Height          =   420
      Left            =   4950
      Max             =   32766
      Min             =   1
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Value           =   15000
      Width           =   270
   End
   Begin VB.ComboBox cboTarget 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   1800
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   1800
      Width           =   1125
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   480
      TabIndex        =   11
      Top             =   1065
      Width           =   555
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   3360
      TabIndex        =   10
      Top             =   1065
      Width           =   600
   End
   Begin VB.Label lblWidthType 
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2490
      TabIndex        =   9
      Top             =   1065
      Width           =   855
   End
   Begin VB.Label lblHeightType 
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   5370
      TabIndex        =   8
      Top             =   1065
      Width           =   855
   End
   Begin VB.Label lblAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Render tiled image using:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   360
      TabIndex        =   2
      Top             =   405
      Width           =   2070
   End
End
Attribute VB_Name = "FormTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tile Rendering Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 25/August/12
'Last updated: 25/August/12
'Last update: Initial build
'
'Render tiled images.  Options are provided for rendering to current wallpaper size, or to a custom size.
' The interface options for custom size are derived from the Image Size form; ideally, any changes to that
' should be mirrored here.
'
'***************************************************************************

Option Explicit

'Used to prevent the scroll bars from getting stuck in update loops
Dim updateWidthBar As Boolean, updateHeightBar As Boolean

'Track the last type of option used; we use this to convert the text box values intelligently
Dim lastTargetMode As Long

'When the combo box is changed, make the appropriate controls visible
Private Sub cboTarget_Click()

    Select Case cboTarget.ListIndex
        'Wallpaper size
        Case 0
            
            'Determine the current screen size, in pixels; this is used to provide a "render to screen size" option
            Dim cScreenWidth As Long, cScreenHeight As Long
            cScreenWidth = Screen.Width / Screen.TwipsPerPixelX
            cScreenHeight = Screen.Height / Screen.TwipsPerPixelY
            
            'Add one to the displayed width and height, since we store them -1 for loops
            txtWidth.Text = cScreenWidth
            txtHeight.Text = cScreenHeight
            
            txtWidth.Enabled = False
            txtHeight.Enabled = False
            VSWidth.Enabled = False
            VSHeight.Enabled = False
            lblWidthType = "pixels"
            lblHeightType = "pixels"
        
        'Custom size (in pixels)
        Case 1
            txtWidth.Enabled = True
            txtHeight.Enabled = True
            VSWidth.Enabled = True
            VSHeight.Enabled = True
            lblWidthType = "pixels"
            lblHeightType = "pixels"
            
            'If the user was previously measuring in tiles, convert that value to pixels
            If (lastTargetMode = 2) And (NumberValid(txtWidth)) And (NumberValid(txtHeight)) Then
                GetImageData
                txtWidth = (val(txtWidth) * (PicWidthL + 1)) - 1
                txtHeight = (val(txtHeight) * (PicHeightL + 1)) - 1
            End If
            
        'Custom size (as number of tiles)
        Case 2
            txtWidth.Enabled = True
            txtHeight.Enabled = True
            VSWidth.Enabled = True
            VSHeight.Enabled = True
            lblWidthType = "tiles"
            lblHeightType = "tiles"
            
            'Since the user will have previously been measuring in pixels, convert that value to tiles
            If NumberValid(txtWidth) And NumberValid(txtHeight) Then
                txtWidth = CLng(CSng(txtWidth) / (PicWidthL + 1))
                txtHeight = CLng(CSng(txtHeight) / (PicHeightL + 1))
            End If
    End Select
    
    'Remember this value for future conversions
    lastTargetMode = cboTarget.ListIndex

End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()

    'Before rendering anything, check to make sure the text boxes have valid input
    If Not EntryValid(txtWidth, 1, 32767, True, True) Then
        AutoSelectText txtWidth
        Exit Sub
    End If
    If Not EntryValid(txtHeight, 1, 32767, True, True) Then
        AutoSelectText txtHeight
        Exit Sub
    End If

    Me.Visible = False
    
    'Based on the user's selection, submit the proper processor request
    Process Tile, cboTarget.ListIndex, txtWidth, txtHeight
    
    Unload Me
    
End Sub

'This routine renders the current image to a new, tiled image (larger than the present image)
' tType is the parameter used for determining how many tiles to draw:
' 0 - current wallpaper size
' 1 - custom size, in pixels
' 2 - custom size, as number of tiles
' The other two parameters are width and height, or tiles in x and y direction
Public Sub GenerateTile(ByVal tType As Byte, Optional xTarget As Long, Optional yTarget As Long)
    
    Message "Rendering tiled image..."
    
    GetImageData
    
    'We need to determine a target width and height based on the input parameters
    Dim targetWidth As Long, targetHeight As Long
    
    PicWidthL = PicWidthL + 1
    PicHeightL = PicHeightL + 1
    
    Select Case tType
        Case 0
            'Current wallpaper size
            targetWidth = Screen.Width / Screen.TwipsPerPixelX
            targetHeight = Screen.Height / Screen.TwipsPerPixelY
        Case 1
            'Custom size
            targetWidth = xTarget
            targetHeight = yTarget
        Case 2
            'Specific number of tiles; determine the target size in pixels, accordingly
            targetWidth = (PicWidthL * xTarget)
            targetHeight = (PicHeightL * yTarget)
    End Select
    
    'Make sure the target width/height isn't too large
    If targetWidth > 32768 Then targetWidth = 32768
    If targetHeight > 32768 Then targetHeight = 32768
    
    'Resize the target picture box to this new size
    FormMain.ActiveForm.BackBuffer2.Width = targetWidth + 2
    FormMain.ActiveForm.BackBuffer2.Height = targetHeight + 2
    
    'Figure out how many loop intervals we'll need in the x and y direction to fill the target size
    Dim xLoop As Long, yLoop As Long
    xLoop = CLng(CSng(targetWidth) / CSng(PicWidthL))
    yLoop = CLng(CSng(targetHeight) / CSng(PicHeightL))
    
    SetProgBarMax xLoop
    
    'Using that loop variable, render the original image to the target picture box that many times
    For x = 0 To xLoop
    For y = 0 To yLoop
        BitBlt FormMain.ActiveForm.BackBuffer2.hDC, x * PicWidthL, y * PicHeightL, PicWidthL, PicHeightL, FormMain.ActiveForm.BackBuffer.hDC, 0, 0, vbSrcCopy
    Next y
        SetProgBarVal x
    Next x
    
    SetProgBarVal xLoop
    
    'With that complete, copy the target back into the original picture box
    FormMain.ActiveForm.BackBuffer.Width = FormMain.ActiveForm.BackBuffer2.Width
    FormMain.ActiveForm.BackBuffer.Height = FormMain.ActiveForm.BackBuffer2.Height
    FormMain.ActiveForm.Picture = LoadPicture("")
    BitBlt FormMain.ActiveForm.BackBuffer.hDC, 0, 0, FormMain.ActiveForm.BackBuffer.ScaleWidth, FormMain.ActiveForm.BackBuffer.ScaleHeight, FormMain.ActiveForm.BackBuffer2.hDC, 0, 0, vbSrcCopy
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    
    'Clear out the secondary picture box to save on memory
    FormMain.ActiveForm.BackBuffer2.Picture = LoadPicture("")
    FormMain.ActiveForm.BackBuffer2.Width = 1
    FormMain.ActiveForm.BackBuffer2.Height = 1
    
    'Display the new size
    DisplaySize FormMain.ActiveForm.BackBuffer.ScaleWidth, FormMain.ActiveForm.BackBuffer.ScaleHeight
    
    SetProgBarVal 0
    
    'Render it on-screen at an automatically set zoom
    FitOnScreen
    
    Message "Finished."

End Sub

'LOAD form
Private Sub Form_Load()
        
    cboTarget.AddItem "Current screen size", 0
    cboTarget.AddItem "Custom size (in pixels)", 1
    cboTarget.AddItem "Specific number of tiles", 2
    cboTarget.ListIndex = 0
    DoEvents
    
    'Create the image previews
    'DrawPreviewImage PicPreview
    'DrawPreviewImage PicEffect
    'PreviewTile 1
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'When the text boxes are changed, keep the scroll bar values in sync
Private Sub txtHeight_Change()
    If EntryValid(txtHeight, 1, 32767, False, True) Then
        updateHeightBar = False
        VSHeight.Value = Abs(32767 - CInt(txtHeight))
        updateHeightBar = True
    Else
        AutoSelectText txtHeight
    End If
End Sub

Private Sub txtHeight_GotFocus()
    AutoSelectText txtHeight
End Sub

Private Sub txtWidth_Change()
    If EntryValid(txtWidth, 1, 32767, False, True) Then
        updateWidthBar = False
        VSWidth.Value = Abs(32767 - CInt(txtWidth))
        updateWidthBar = True
    Else
        AutoSelectText txtWidth
    End If
End Sub

Private Sub txtWidth_GotFocus()
    AutoSelectText txtWidth
End Sub

'When the scroll bars are changed, keep the text box values in sync
Private Sub VSHeight_Change()
    If updateHeightBar = True Then txtHeight = Abs(32767 - CStr(VSHeight.Value))
End Sub

Private Sub VSWidth_Change()
    If updateWidthBar = True Then txtWidth = Abs(32767 - CStr(VSWidth.Value))
End Sub
