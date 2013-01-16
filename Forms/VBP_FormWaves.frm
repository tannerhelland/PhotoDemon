VERSION 5.00
Begin VB.Form FormWaves 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Waves"
   ClientHeight    =   6540
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   12090
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbEdges 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3975
      Width           =   4860
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10590
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin VB.TextBox txtWavelengthY 
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
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "1"
      Top             =   2220
      Width           =   735
   End
   Begin VB.HScrollBar hsWavelengthY 
      Height          =   255
      Left            =   6120
      Max             =   50
      Min             =   1
      TabIndex        =   13
      Top             =   2280
      Value           =   1
      Width           =   4815
   End
   Begin VB.TextBox txtAmplitudeY 
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
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "0"
      Top             =   3060
      Width           =   735
   End
   Begin VB.HScrollBar hsAmplitudeY 
      Height          =   255
      Left            =   6120
      Max             =   200
      TabIndex        =   11
      Top             =   3120
      Width           =   4815
   End
   Begin VB.HScrollBar hsAmplitudeX 
      Height          =   255
      Left            =   6120
      Max             =   200
      TabIndex        =   9
      Top             =   1440
      Value           =   20
      Width           =   4815
   End
   Begin VB.TextBox txtAmplitudeX 
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
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "20"
      Top             =   1380
      Width           =   735
   End
   Begin VB.HScrollBar hsWavelengthX 
      Height          =   255
      Left            =   6120
      Max             =   50
      Min             =   1
      TabIndex        =   6
      Top             =   600
      Value           =   30
      Width           =   4815
   End
   Begin VB.TextBox txtWavelengthX 
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
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "30"
      Top             =   540
      Width           =   735
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
      Left            =   6120
      TabIndex        =   4
      Top             =   4860
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
      Left            =   7560
      TabIndex        =   3
      Top             =   4860
      Width           =   2535
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   5625
      _extentx        =   9922
      _extenty        =   9922
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "if pixels lie outside the image..."
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
      Index           =   5
      Left            =   6000
      TabIndex        =   19
      Top             =   3600
      Width           =   3315
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   17
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vertical wavelength:"
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
      Index           =   4
      Left            =   6000
      TabIndex        =   16
      Top             =   1920
      Width           =   2115
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vertical strength (amplitude):"
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
      Left            =   6000
      TabIndex        =   15
      Top             =   2760
      Width           =   3090
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "horizontal strength (amplitude):"
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
      Index           =   3
      Left            =   6000
      TabIndex        =   10
      Top             =   1080
      Width           =   3390
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "horizontal wavelength:"
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
      Left            =   6000
      TabIndex        =   7
      Top             =   240
      Width           =   2415
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
      Left            =   6000
      TabIndex        =   2
      Top             =   4500
      Width           =   1845
   End
End
Attribute VB_Name = "FormWaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image "Waves" Distortion
'Copyright ©2000-2013 by Tanner Helland
'Created: 07/January/13
'Last updated: 15/January/13
'Last update: added user-settable options for edge handling
'
'This tool allows the user to apply a "waves" distortion to an image.  Bilinear interpolation
' (via reverse-mapping) is available for a high-quality result.
'
'Four parameters are required - wavelength and amplitude for both the x and y directions.  Unlike other distortions
' in PhotoDemon, this one doesn't require a radius.
'
'Finally, the transformation used by this tool is a modified version of a transformation originally written by
' Jerry Huxtable of JH Labs.  Jerry's original code is licensed under an Apache 2.0 license.  You may download his
' original version at the following link (good as of 07 January '13): http://www.jhlabs.com/ip/filters/index.html
'
'***************************************************************************

Option Explicit

'Use this to prevent the text box and scroll bar from updating each other in an endless loop
Dim userChange As Boolean

Private Sub cmbEdges_Click()
    updatePreview
End Sub

Private Sub cmbEdges_Scroll()
    updatePreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'Before rendering anything, check to make sure the text boxes have valid input
    If Not EntryValid(txtWavelengthX, hsWavelengthX.Min, hsWavelengthX.Max, True, True) Then
        AutoSelectText txtWavelengthX
        Exit Sub
    End If
    
    If Not EntryValid(txtAmplitudeX, hsAmplitudeX.Min, hsAmplitudeX.Max, True, True) Then
        AutoSelectText txtAmplitudeX
        Exit Sub
    End If
    
    If Not EntryValid(txtWavelengthY, hsWavelengthY.Min, hsWavelengthY.Max, True, True) Then
        AutoSelectText txtWavelengthY
        Exit Sub
    End If
    
    If Not EntryValid(txtAmplitudeY, hsAmplitudeY.Min, hsAmplitudeY.Max, True, True) Then
        AutoSelectText txtAmplitudeY
        Exit Sub
    End If

    Me.Visible = False
    
    'Based on the user's selection, submit the proper processor request
    Process DistortWaves, CDbl(hsWavelengthX), CDbl(hsAmplitudeX), CDbl(hsWavelengthY), CDbl(hsAmplitudeY), CLng(cmbEdges.ListIndex), OptInterpolate(0)
        
    Unload Me
    
End Sub

'Apply a "wave-like" effect to an image
Public Sub WaveImage(ByVal xWavelength As Double, ByVal xAmplitude As Double, ByVal yWavelength As Double, ByVal yAmplitude As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If toPreview = False Then Message "Dipping image in virtual wave pool..."
        
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
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickVal2 As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, curLayerValues.MaxX, curLayerValues.MaxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
          
    'This wave transformation requires specialized variables
    xWavelength = 51 - xWavelength
    yWavelength = 51 - yWavelength
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
    
    Dim i As Long
                 
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Remap the coordinates around a center point of (0, 0)
        nX = (y - yAmplitude) / xWavelength
        nY = (x - xAmplitude) / yWavelength
        
        srcX = x + Sin(nX) * xAmplitude
        srcY = y + Sin(nY) * yAmplitude
                
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                        
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
        
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    popDistortEdgeBox cmbEdges, EDGE_REFLECT
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Mark scroll bar changes as coming from the user
    userChange = True
    
    'Create the preview
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub hsAmplitudeX_Change()
    copyToTextBoxI txtAmplitudeX, hsAmplitudeX.Value
    updatePreview
End Sub

Private Sub hsAmplitudeX_Scroll()
    copyToTextBoxI txtAmplitudeX, hsAmplitudeX.Value
    updatePreview
End Sub

Private Sub hsWavelengthX_Change()
    copyToTextBoxI txtWavelengthX, hsWavelengthX.Value
    updatePreview
End Sub

Private Sub hsWavelengthX_Scroll()
    copyToTextBoxI txtWavelengthX, hsWavelengthX.Value
    updatePreview
End Sub

Private Sub hsAmplitudeY_Change()
    copyToTextBoxI txtAmplitudeY, hsAmplitudeY.Value
    updatePreview
End Sub

Private Sub hsAmplitudeY_Scroll()
    copyToTextBoxI txtAmplitudeY, hsAmplitudeY.Value
    updatePreview
End Sub

Private Sub hsWavelengthY_Change()
    copyToTextBoxI txtWavelengthY, hsWavelengthY.Value
    updatePreview
End Sub

Private Sub hsWavelengthY_Scroll()
    copyToTextBoxI txtWavelengthY, hsWavelengthY.Value
    updatePreview
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub txtAmplitudeX_GotFocus()
    AutoSelectText txtAmplitudeX
End Sub

Private Sub txtAmplitudeX_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtAmplitudeX
    If EntryValid(txtAmplitudeX, hsAmplitudeX.Min, hsAmplitudeX.Max, False, False) Then hsAmplitudeX.Value = Val(txtAmplitudeX)
End Sub

Private Sub txtWavelengthX_GotFocus()
    AutoSelectText txtWavelengthX
End Sub

Private Sub txtWavelengthX_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtWavelengthX
    If EntryValid(txtWavelengthX, hsWavelengthX.Min, hsWavelengthX.Max, False, False) Then hsWavelengthX.Value = Val(txtWavelengthX)
End Sub

Private Sub txtAmplitudeY_GotFocus()
    AutoSelectText txtAmplitudeY
End Sub

Private Sub txtAmplitudeY_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtAmplitudeY
    If EntryValid(txtAmplitudeY, hsAmplitudeY.Min, hsAmplitudeY.Max, False, False) Then hsAmplitudeY.Value = Val(txtAmplitudeY)
End Sub

Private Sub txtWavelengthY_GotFocus()
    AutoSelectText txtWavelengthY
End Sub

Private Sub txtWavelengthY_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtWavelengthY
    If EntryValid(txtWavelengthY, hsWavelengthY.Min, hsWavelengthY.Max, False, False) Then hsWavelengthY.Value = Val(txtWavelengthY)
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()

    WaveImage CDbl(hsWavelengthX), CDbl(hsAmplitudeX), CDbl(hsWavelengthY), CDbl(hsAmplitudeY), CLng(cmbEdges.ListIndex), OptInterpolate(0), True, fxPreview
    
End Sub
