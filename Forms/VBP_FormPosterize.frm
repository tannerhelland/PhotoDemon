VERSION 5.00
Begin VB.Form FormPosterize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Posterize"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6255
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
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsBits 
      Height          =   255
      Left            =   360
      Max             =   7
      Min             =   1
      TabIndex        =   1
      Top             =   3840
      Value           =   7
      Width           =   4935
   End
   Begin VB.TextBox txtBits 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   5400
      TabIndex        =   0
      Text            =   "7"
      Top             =   3810
      Width           =   495
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   4680
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   4680
      Width           =   1125
   End
   Begin VB.Label lblBeforeandAfter 
      BackStyle       =   0  'Transparent
      Caption         =   "  Before                                                           After"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "# of Bits:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   765
   End
End
Attribute VB_Name = "FormPosterize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Posterizing Effect Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 6/August/06
'Last update: previewing, optimization, comments, variable type changes
'
'Updated posterizing interface; it has been optimized for speed and
'  ease-of-implementation.  If only VB had bit-shift operators....
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    If EntryValid(txtBits, hsBits.Min, hsBits.Max) Then
        Me.Visible = False
        Process Posterize, hsBits.Value
        Unload Me
    Else
        AutoSelectText txtBits
    End If
End Sub

'Subroutine for reducing the representative bits in an image
Public Sub PosterizeImage(ByVal NumOfBits As Byte, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Posterizing image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'pStep is the distance between values that X number of bits allows
    Dim pStep As Double
    pStep = 255 / (2 ^ CLng(NumOfBits) - 1)
    
    'Look-up tables make this far more efficient
    Dim LookUp(0 To 255) As Byte
    For x = 0 To 255
        'Add 0.5 so that values are rounded, not truncated (slightly better results)
        LookUp(x) = CByte(Int(Int(CDbl(x) / pStep + 0.5) * pStep))
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'The posterize can be implemented in a single look-up per color
        ImageData(QuickVal + 2, y) = LookUp(ImageData(QuickVal + 2, y))
        ImageData(QuickVal + 1, y) = LookUp(ImageData(QuickVal + 1, y))
        ImageData(QuickVal, y) = LookUp(ImageData(QuickVal, y))
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
     
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Create the previews
    DrawPreviewImage picPreview
    PosterizeImage hsBits.Value, True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'The following routines are for keeping the text box and scroll bar values in lock-step
Private Sub hsBits_Change()
    txtBits.Text = hsBits.Value
    PosterizeImage hsBits.Value, True, picEffect
End Sub

Private Sub hsBits_Scroll()
    txtBits.Text = hsBits.Value
    PosterizeImage hsBits.Value, True, picEffect
End Sub

Private Sub txtBits_Change()
    If EntryValid(txtBits, hsBits.Min, hsBits.Max, False, False) Then hsBits.Value = val(txtBits)
End Sub

Private Sub txtBits_GotFocus()
    AutoSelectText txtBits
End Sub
