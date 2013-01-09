VERSION 5.00
Begin VB.Form FormNoise 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Add Noise"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
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
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsNoise 
      Height          =   255
      Left            =   360
      Max             =   500
      Min             =   1
      TabIndex        =   3
      Top             =   3840
      Value           =   1
      Width           =   4935
   End
   Begin VB.TextBox txtNoise 
      Alignment       =   2  'Center
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
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "1"
      Top             =   3780
      Width           =   615
   End
   Begin VB.CheckBox ChkM 
      Appearance      =   0  'Flat
      Caption         =   "monochromatic noise only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   465
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   5040
      Width           =   1245
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label lblAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "after"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3360
      TabIndex        =   9
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblBefore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "before"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   900
   End
End
Attribute VB_Name = "FormNoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Noise Interface
'Copyright ©2000-2013 by Tanner Helland
'Created: 3/15/01
'Last updated: 10/September/12
'Last update: rewrote against new layer class
'
'Form for adding noise to an image.
'
'***************************************************************************

Option Explicit

Private Sub ChkM_Click()
    AddNoise hsNoise.Value, ChkM.Value, True, picEffect
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    If EntryValid(txtNoise, hsNoise.Min, hsNoise.Max) Then
        FormNoise.Visible = False
        Process Noise, hsNoise.Value, ChkM.Value
        Unload Me
    Else
        AutoSelectText txtNoise
    End If
End Sub

'Subroutine for adding noise to an image
' Inputs: Amount of noise, monochromatic or not, preview settings
Public Sub AddNoise(ByVal Noise As Long, ByVal MC As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Increasing image noise..."
    
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
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Noise variables
    Dim nColor As Long
    Dim dNoise As Long
    
    'Double the amount of noise we plan on using (so we can add noise above or below the current color value)
    dNoise = Noise * 2
    
    'Although it's slow, we're stuck using random numbers for noise addition.  Seed the generator with a pseudo-random value.
    Randomize Timer
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        If MC = True Then
            
            'Monochromatic noise - same amount for each color
            nColor = (dNoise * Rnd) - Noise
            r = r + nColor
            g = g + nColor
            b = b + nColor
            
        Else
            
            'Colored noise - each color generated randomly
            r = r + (dNoise * Rnd) - Noise
            g = g + (dNoise * Rnd) - Noise
            b = b + (dNoise * Rnd) - Noise
            
        End If
        
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        
        'Assign that blended value to each color channel
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
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

Private Sub Form_Activate()

    'Create the previews
    DrawPreviewImage picPreview
    AddNoise hsNoise.Value, ChkM.Value, True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The following four routines keep the value of the textbox and scroll bar in lock-step
Private Sub hsNoise_Change()
    copyToTextBoxI txtNoise, hsNoise.Value
    AddNoise hsNoise.Value, ChkM.Value, True, picEffect
End Sub

Private Sub hsNoise_Scroll()
    copyToTextBoxI txtNoise, hsNoise.Value
    AddNoise hsNoise.Value, ChkM.Value, True, picEffect
End Sub

Private Sub txtNoise_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtNoise
    If EntryValid(txtNoise, hsNoise.Min, hsNoise.Max, False, False) Then hsNoise.Value = Val(txtNoise)
End Sub

Private Sub txtNoise_GotFocus()
    AutoSelectText txtNoise
End Sub
