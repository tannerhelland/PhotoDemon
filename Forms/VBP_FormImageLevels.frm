VERSION 5.00
Begin VB.Form FormImageLevels 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Adjust Image Levels"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   195
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
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsOutL 
      Height          =   220
      Left            =   1200
      Max             =   255
      TabIndex        =   3
      Top             =   4800
      Width           =   4455
   End
   Begin VB.HScrollBar hsOutR 
      Height          =   220
      Left            =   1200
      Max             =   255
      TabIndex        =   4
      Top             =   5040
      Value           =   255
      Width           =   4455
   End
   Begin VB.HScrollBar hsInR 
      Height          =   220
      Left            =   1200
      Max             =   255
      Min             =   2
      TabIndex        =   2
      Top             =   4080
      Value           =   255
      Width           =   4455
   End
   Begin VB.HScrollBar hsInL 
      Height          =   220
      Left            =   1200
      Max             =   253
      TabIndex        =   0
      Top             =   3600
      Width           =   4455
   End
   Begin VB.HScrollBar hsInM 
      Height          =   220
      Left            =   1200
      Max             =   254
      Min             =   1
      TabIndex        =   1
      Top             =   3840
      Value           =   127
      Width           =   4455
   End
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      Caption         =   "Reset scrollbars"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   5880
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   5880
      Width           =   1125
   End
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
      TabIndex        =   8
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
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblOutputLevels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Output levels:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   4560
      Width           =   5775
   End
   Begin VB.Label lblOutputL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left limit:    0"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   4800
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5745
      TabIndex        =   22
      Top             =   4800
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right limit:  0"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   5040
      Width           =   930
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5745
      TabIndex        =   20
      Top             =   5040
      Width           =   270
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Input levels:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3360
      Width           =   5775
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5745
      TabIndex        =   18
      Top             =   4080
      Width           =   270
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right limit:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   4080
      Width           =   750
   End
   Begin VB.Label lblLeftR 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "253"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5745
      TabIndex        =   16
      Top             =   3600
      Width           =   270
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left limit:    0"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   930
   End
   Begin VB.Label lblMiddleR 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "254"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5745
      TabIndex        =   14
      Top             =   3840
      Width           =   270
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Midtones:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   705
   End
   Begin VB.Label lblMiddleL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1080
      TabIndex        =   12
      Top             =   3840
      Width           =   90
   End
   Begin VB.Label lblRightL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1080
      TabIndex        =   11
      Top             =   4080
      Width           =   90
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
      TabIndex        =   10
      Top             =   2880
      Width           =   3975
   End
End
Attribute VB_Name = "FormImageLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Levels
'Copyright ©2006-2012 by Tanner Helland
'Created: 22/July/06
'Last updated: 09/September/12
'Last update: rewrote everything against the new layer class
'
'This form is an exact model of how to adjust image levels (identical to
' Photoshop's method).  Be forewarned - there are some fairly involved (i.e. incomprehensible)
' math sections.
'
'***************************************************************************

Option Explicit

'Constants required for creating a gamma curve from .1 to 10
Private Const MAXGAMMA As Double = 1.8460498941512
Private Const MIDGAMMA As Double = 0.68377223398334
Private Const ROOT10 As Double = 3.16227766

'Used to track the ratio of the midtones scrollbar, so that when the left and
'right values get changed, we automatically set the midtone to the same ratio
'(i.e. as Photoshop does it)
Dim midRatio As Double

'Whether or not changing the midtone scrollbar is user-generated or program-generated
'(so we only refresh if the user moved it - otherwise we get bad looping)
Dim iRefresh As Boolean

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    
    Me.Visible = False
    
    Process ImageLevels, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
    
    Unload Me
    
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Draw preview images to the top picture boxes
    DrawPreviewImage picPreview
    MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
    
    'Set the default midtone scrollbar ratio to 1/2
    midRatio = 0.5
    
    '...and allow refreshing
    iRefresh = True
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'This will reset the scrollbars to default levels
Private Sub cmdReset_Click()
    
    'Allow refreshing
    iRefresh = True
    
    'Set the output levels to (0-255)
    hsOutL.Value = 0
    hsOutR.Value = 255
    
    'Set the input levels to (0-255)
    hsInL.Value = 0
    hsInR.Value = 255
    FixScrollBars
    
    'Set the midtone level to default (127)
    midRatio = 0.5
    hsInM.Value = 127
    FixScrollBars
    
End Sub


'*********************************************************************************
'The following 10 subroutines are for changing/scrolling any of the scrollbars
'on the main form
'*********************************************************************************
Private Sub hsInL_Change()
    FixScrollBars
    MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
End Sub

Private Sub hsInL_Scroll()
    FixScrollBars
    MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
End Sub

Private Sub hsInM_Change()
    If iRefresh = True Then
        midRatio = (CDbl(hsInM.Value) - CDbl(hsInL.Value)) / (CDbl(hsInR.Value) - CDbl(hsInL.Value))
        FixScrollBars True
        MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
    End If
End Sub

Private Sub hsInM_Scroll()
    If iRefresh = True Then
        midRatio = (CDbl(hsInM.Value) - CDbl(hsInL.Value)) / (CDbl(hsInR.Value) - CDbl(hsInL.Value))
        FixScrollBars True
        MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
    End If
End Sub

Private Sub hsInR_Change()
    FixScrollBars
    MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
End Sub

Private Sub hsInR_Scroll()
    FixScrollBars
    MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
End Sub

Private Sub hsOutL_Change()
    MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
End Sub

Private Sub hsOutL_Scroll()
    MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
End Sub

Private Sub hsOutR_Change()
    MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
End Sub

Private Sub hsOutR_Scroll()
    MapImageLevels hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value, True, PicEffect
End Sub


'Draw an image based on user-adjusted input and output levels
Public Sub MapImageLevels(ByVal inLLimit As Long, ByVal inMLimit As Long, ByVal inRLimit As Long, ByVal outLLimit As Long, ByVal outRLimit As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Mapping new image levels..."
    
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
        
    'Look-up table for the midtone (gamma) leveled values
    Dim gValues(0 To 255) As Double
    
    'WARNING: This next chunk of code is a lot of messy math.  Don't worry too much
    ' if you can't make sense of it ;)
    
    'Fill the gamma table with appropriate gamma values (from 10 to .1, ranged quadratically)
    ' NOTE: This table is constant, and could theoretically be loaded from file instead of generated
    ' every time we run this function.
    Dim gStep As Double
    gStep = (MAXGAMMA + MIDGAMMA) / 127
    For x = 0 To 127
        gValues(x) = (CDbl(x) / 127) * MIDGAMMA
    Next x
    For x = 128 To 255
        gValues(x) = MIDGAMMA + (CDbl(x - 127) * gStep)
    Next x
    For x = 0 To 255
        gValues(x) = 1 / ((gValues(x) + 1 / ROOT10) ^ 2)
    Next x
    
    'Because we've built our look-up tables on a 0-255 scale, correct the inMLimit
    ' value (from the midtones scroll bar) to simply represent a ratio on that scale
    Dim tRatio As Double
    tRatio = (inMLimit - inLLimit) / (inRLimit - inLLimit)
    tRatio = tRatio * 255
    
    'Then convert that ratio into a byte (so we can access a look-up table with it)
    Dim bRatio As Byte
    bRatio = CByte(tRatio)
    
    'Calculate a look-up table of gamma-corrected values based on the midtones scrollbar
    Dim gLevels(0 To 255) As Byte
    Dim tmpGamma As Double
    For x = 0 To 255
        tmpGamma = CDbl(x) / 255
        tmpGamma = tmpGamma ^ (1 / gValues(bRatio))
        tmpGamma = tmpGamma * 255
        If tmpGamma > 255 Then
            tmpGamma = 255
        ElseIf tmpGamma < 0 Then
            tmpGamma = 0
        End If
        gLevels(x) = tmpGamma
    Next x
    
    'Look-up table for the input leveled values
    Dim newLevels(0 To 255) As Byte
    
    'Fill the look-up table with appropriately mapped input limits
    Dim pStep As Single
    pStep = 255 / (CSng(inRLimit) - CSng(inLLimit))
    For x = 0 To 255
        If x < inLLimit Then
            newLevels(x) = 0
        ElseIf x > inRLimit Then
            newLevels(x) = 255
        Else
            newLevels(x) = ByteMe(((CSng(x) - CSng(inLLimit)) * pStep))
        End If
    Next x
    
    'Now run all input-mapped values through our midtone-correction look-up
    For x = 0 To 255
        newLevels(x) = gLevels(newLevels(x))
    Next x
    
    'Last of all, remap all image values to match the user-specified output limits
    Dim oStep As Double
    oStep = (CSng(outRLimit) - CSng(outLLimit)) / 255
    For x = 0 To 255
        newLevels(x) = ByteMe(CSng(outLLimit) + (CSng(newLevels(x)) * oStep))
    Next x
    
    'Now we can finally loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Assign new values looking the lookup table
        ImageData(QuickVal + 2, y) = newLevels(r)
        ImageData(QuickVal + 1, y) = newLevels(g)
        ImageData(QuickVal, y) = newLevels(b)
        
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

'Used to make sure the scroll bars have appropriate limits
Private Sub FixScrollBars(Optional midMoving As Boolean = False)

    'Make sure that the input scrollbar values don't overlap, and update the labels
    'to display such
    hsInM.Min = hsInL.Value + 1
    lblMiddleL.Caption = hsInL.Value + 1
    hsInR.Min = hsInL.Value + 2
    lblRightL.Caption = hsInL.Value + 2
    hsInL.Max = hsInR.Value - 2
    lblLeftR.Caption = hsInR.Value - 2
    hsInM.Max = hsInR.Value - 1
    lblMiddleR.Caption = hsInR.Value - 1
    
    'If the user hasn't moved the midtones scrollbar, attempt to preserve its ratio
    If midMoving = False Then
        iRefresh = False
        Dim newValue As Long
        newValue = hsInL.Value + midRatio * (CDbl(hsInR.Value) - CDbl(hsInL.Value))
        If newValue > hsInM.Max Then
            newValue = hsInM.Max
        ElseIf newValue < hsInM.Min Then
            newValue = hsInM.Min
        End If
        hsInM.Value = newValue
        DoEvents
        iRefresh = True
    End If
    
End Sub

'Used to convert Long-type variables to bytes (with proper [0,255] range)
Private Function ByteMe(ByVal val As Long) As Byte
    If val > 255 Then
        ByteMe = 255
    ElseIf val < 0 Then
        ByteMe = 0
    Else
        ByteMe = val
    End If
End Function

