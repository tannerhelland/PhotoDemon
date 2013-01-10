VERSION 5.00
Begin VB.Form FormFiguredGlass 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Figured Glass"
   ClientHeight    =   9180
   ClientLeft      =   -15
   ClientTop       =   225
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
   ScaleHeight     =   612
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   8550
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4710
      TabIndex        =   1
      Top             =   8550
      Width           =   1365
   End
   Begin VB.HScrollBar hsScale 
      Height          =   255
      Left            =   360
      Max             =   100
      Min             =   1
      TabIndex        =   11
      Top             =   6060
      Value           =   50
      Width           =   4815
   End
   Begin VB.TextBox txtScale 
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
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "50"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox txtTurbulence 
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
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "50"
      Top             =   6780
      Width           =   735
   End
   Begin VB.HScrollBar hsTurbulence 
      Height          =   255
      Left            =   360
      Max             =   100
      Min             =   1
      TabIndex        =   7
      Top             =   6840
      Value           =   50
      Width           =   4815
   End
   Begin VB.OptionButton OptInterpolate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " quality"
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
      Height          =   360
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   7680
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton OptInterpolate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " speed"
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
      Height          =   360
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   7680
      Width           =   2535
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5100
      Left            =   240
      ScaleHeight     =   338
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   382
      TabIndex        =   3
      Top             =   240
      Width           =   5760
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -840
      TabIndex        =   12
      Top             =   8400
      Width           =   7095
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "turbulence:"
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
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   6480
      Width           =   1200
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "render emphasis:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   7320
      Width           =   1845
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "scale:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   600
   End
End
Attribute VB_Name = "FormFiguredGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image "Figured Glass" Distortion
'Copyright ©2000-2013 by Tanner Helland
'Created: 08/January/13
'Last updated: 08/January/13
'Last update: initial build
'
'This tool allows the user to apply a distort operation to an image that mimicks seeing it through warped glass, perhaps
' glass tiles of some sort.  Many different names are used for this effect - Paint.NET calls it "dents" (which I quite
' dislike); other software calls it "marbling".  I chose figured glass because it's an actual type of uneven glass - see:
' http://en.wikipedia.org/wiki/Architectural_glass#Rolled_plate_.28figured.29_glass
'
'As with other distorts in the program, bilinear interpolation (via reverse-mapping) is available for a
' high-quality transformation.
'
'No radius is required for the effect.  It always operates on the entire image.
'
'Finally, the transformation used by this tool is a modified version of a transformation originally written by
' Jerry Huxtable of JH Labs.  Jerry's original code is licensed under an Apache 2.0 license.  You may download his
' original version at the following link (good as of 07 January '13): http://www.jhlabs.com/ip/filters/index.html
'
'***************************************************************************

Option Explicit

'Use this to prevent the text box and scroll bar from updating each other in an endless loop
Dim userChange As Boolean

'This variable stores random z-location in the perlin noise generator (which allows for a unique effect each time the form is loaded)
Dim m_zOffset As Double

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'Before rendering anything, check to make sure the text boxes have valid input
    If Not EntryValid(txtTurbulence, hsTurbulence.Min, hsTurbulence.Max, True, True) Then
        AutoSelectText txtTurbulence
        Exit Sub
    End If

    If Not EntryValid(txtScale, hsScale.Min, hsScale.Max, True, True) Then
        AutoSelectText txtScale
        Exit Sub
    End If
    
    Me.Visible = False
    
    'Based on the user's selection, submit the proper processor request
    If OptInterpolate(0) Then
        Process DistortFiguredGlass, CDbl(hsScale), CDbl(hsTurbulence / 100), True
    Else
        Process DistortFiguredGlass, CDbl(hsScale), CDbl(hsTurbulence / 100), False
    End If
    
    Unload Me
    
End Sub

'Apply a "figured glass" effect to an image
Public Sub FiguredGlassFX(ByVal fxScale As Double, ByVal fxTurbulence As Double, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Projecting image through simulated glass..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'Because interpolation may be used, it's necessary to keep pixel values within special ranges
    Dim xLimit As Long, yLimit As Long
    If useBilinear Then
        xLimit = finalX - 1
        yLimit = finalY - 1
    Else
        xLimit = finalX
        yLimit = finalY
    End If

    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickVal2 As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
          
    'Our etched glass effect requires some specialized variables
        
    'Invert turbulence
    fxTurbulence = 1.01 - fxTurbulence
        
    'Sin and cosine look-up tables
    Dim sinTable(0 To 255) As Double, cosTable(0 To 255) As Double
    
    'Populate the look-up tables
    Dim fxAngle As Double
    
    Dim i As Long
    For i = 0 To 255
        fxAngle = (PI_DOUBLE * i) / (256 * fxTurbulence)
        sinTable(i) = -fxScale * Sin(fxAngle)
        cosTable(i) = fxScale * Cos(fxAngle)
    Next i
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
                                  
    'This effect requires a noise function to operate.  I use Steve McMahon's excellent Perlin Noise class for this.
    Dim cPerlin As cPerlin3D
    Set cPerlin = New cPerlin3D
        
    'Finally, an integer displacement will be used to move pixel values around
    Dim pDisplace As Long
                                  
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Calculate a displacement for this point
        pDisplace = 127 * (1 + cPerlin.Noise(x / fxScale, y / fxScale, m_zOffset))
        If pDisplace < 0 Then pDisplace = 0
        If pDisplace > 255 Then pDisplace = 255
        
        'Calculate a new source pixel using the sin and cos look-up tables and our calculated displacement
        srcX = x + sinTable(pDisplace)
        srcY = y + sinTable(pDisplace)
        
        'Make sure the source coordinates are in-bounds
        If srcX < 0 Then srcX = 0
        If srcY < 0 Then srcY = 0
        If srcX > xLimit Then srcX = xLimit
        If srcY > yLimit Then srcY = yLimit
        
        'Interpolate the result if desired, otherwise use nearest-neighbor
        If useBilinear Then
                
            For i = 0 To qvDepth - 1
                dstImageData(QuickVal + i, y) = getInterpolatedVal(srcX, srcY, srcImageData, i, qvDepth)
            Next i
        
        Else
        
            QuickVal2 = Int(srcX) * qvDepth
        
            For i = 0 To qvDepth - 1
                dstImageData(QuickVal + i, y) = srcImageData(QuickVal2 + i, Int(srcY))
            Next i
                
        End If
                
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
End Sub

Private Sub Form_Activate()
    
    'Create the preview
    updatePreview
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Mark scroll bar changes as coming from the user
    userChange = True
    
End Sub

Private Sub Form_Load()
    
    'Calculate a random z offset for the noise function
    Randomize Timer
    m_zOffset = Rnd
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Keep the scroll bar and the text box values in sync
Private Sub hsTurbulence_Change()
    copyToTextBoxI txtTurbulence, hsTurbulence.Value
    updatePreview
End Sub

Private Sub hsTurbulence_Scroll()
    copyToTextBoxI txtTurbulence, hsTurbulence.Value
    updatePreview
End Sub

Private Sub hsScale_Scroll()
    copyToTextBoxI txtScale, hsScale.Value
    updatePreview
End Sub

Private Sub hsScale_Change()
    copyToTextBoxI txtScale, hsScale.Value
    updatePreview
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub txtScale_GotFocus()
    AutoSelectText txtScale
End Sub

Private Sub txtScale_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtScale, True
    If EntryValid(txtScale, hsScale.Min, hsScale.Max, False, False) Then hsScale.Value = Val(txtScale)
End Sub

Private Sub txtTurbulence_GotFocus()
    AutoSelectText hsTurbulence
End Sub

Private Sub txtTurbulence_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtTurbulence, True
    If EntryValid(txtTurbulence, hsTurbulence.Min, hsTurbulence.Max, False, False) Then hsTurbulence.Value = Val(txtTurbulence)
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()

    If OptInterpolate(0) Then
        FiguredGlassFX CDbl(hsScale), CDbl(hsTurbulence / 100), True, True, picPreview
    Else
        FiguredGlassFX CDbl(hsScale), CDbl(hsTurbulence / 100), False, True, picPreview
    End If

End Sub
