VERSION 5.00
Begin VB.Form FormLevels 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Adjust Image Levels"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   195
   ClientWidth     =   12180
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   812
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9240
      TabIndex        =   14
      Top             =   5910
      Width           =   1365
   End
   Begin PhotoDemon.sliderTextCombo sltOutL 
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   4050
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      Max             =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10710
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      Caption         =   "&Reset levels"
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   5880
      Width           =   1455
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltOutR 
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   4890
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      Max             =   255
      Value           =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltInL 
      Height          =   495
      Left            =   6240
      TabIndex        =   11
      Top             =   930
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      Max             =   253
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltInR 
      Height          =   495
      Left            =   6240
      TabIndex        =   13
      Top             =   2610
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      Min             =   2
      Max             =   255
      Value           =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltInM 
      Height          =   495
      Left            =   6240
      TabIndex        =   16
      Top             =   1770
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSubHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gray point (midtone):"
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
      Left            =   6240
      TabIndex        =   15
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblSubHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "white point (ceiling):"
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
      Left            =   6240
      TabIndex        =   12
      Top             =   2280
      Width           =   2205
   End
   Begin VB.Label lblSubHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "black point (floor):"
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
      Left            =   6240
      TabIndex        =   10
      Top             =   600
      Width           =   1995
   End
   Begin VB.Label lblSubHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "white point (ceiling):"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   4560
      Width           =   2205
   End
   Begin VB.Label lblSubHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "black point (floor):"
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
      Left            =   6240
      TabIndex        =   6
      Top             =   3720
      Width           =   1995
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   5760
      Width           =   12255
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "output levels"
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
      Left            =   6000
      TabIndex        =   3
      Top             =   3240
      Width           =   1350
   End
   Begin VB.Label lblInput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "input levels"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "FormLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Levels
'Copyright ©2006-2013 by Tanner Helland
'Created: 22/July/06
'Last updated: 26/April/13
'Last update: redesigned the entire form around the new slider/text custom control.  This should make it much more
'              user-friendly, and for the first time ever, values can now be entered via text box.  Yay!
'
'This tool allows the user to adjust image levels.  Its behavior is based off Photoshop's Levels tool, and identical
' values entered into both programs should yield an identical image.
'
'Unfortunately, to perfectly mimic Photoshop's behavior, some fairly involved (i.e. incomprehensible) math is required.
' To mitigate the speed implications of such convoluted math, a number of look-up tables are used.  This makes the
' function quite fast, but at a hit to readability.  My apologies.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Constants required for creating a gamma curve from .1 to 10
Private Const MAXGAMMA As Double = 1.8460498941512
Private Const MIDGAMMA As Double = 0.68377223398334
Private Const ROOT10 As Double = 3.16227766

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    
    Me.Visible = False
    Process ImageLevels, sltInL.Value, sltInM.Value, sltInR.Value, sltOutL.Value, sltOutR.Value
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Draw a preview image
    updatePreview

End Sub

'This will reset the scrollbars to default levels
Private Sub cmdReset_Click()
    
    'Set the output levels to (0-255)
    sltOutL.Value = 0
    sltOutR.Value = 255
    
    'Set the input levels to (0-255)
    sltInL.Value = 0
    sltInR.Value = 255
    FixScrollBars
    
    'Set the midtone level to default (0.5)
    sltInM.Value = 0.5
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Draw an image based on user-adjusted input and output levels
Public Sub MapImageLevels(ByVal inLLimit As Long, ByVal inMLimit As Double, ByVal inRLimit As Long, ByVal outLLimit As Long, ByVal outRLimit As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

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
    
    'Convert the midtone ratio into a byte (so we can access a look-up table with it)
    Dim bRatio As Byte
    bRatio = CByte(inMLimit * 255)
    
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
    Dim pStep As Double
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

'Used to convert Long-type variables to bytes (with proper [0,255] range)
Private Function ByteMe(ByVal bVal As Long) As Byte
    If bVal > 255 Then
        ByteMe = 255
    ElseIf bVal < 0 Then
        ByteMe = 0
    Else
        ByteMe = bVal
    End If
End Function

'Used to make sure the scroll bars have appropriate limits
Private Sub FixScrollBars()
    
    'The black tone input level is never allowed to be > the white tone input level.
    sltInL.Max = sltInR.Value - 2
    
    ' Similarly, the white tone input level is never allowed to be < the black tone input level.
    sltInR.Min = sltInL.Value + 2
    
End Sub

Private Sub sltInL_Change()
    FixScrollBars
    updatePreview
End Sub

Private Sub sltInM_Change()
    updatePreview
End Sub

Private Sub sltInR_Change()
    FixScrollBars
    updatePreview
End Sub

Private Sub sltOutL_Change()
    updatePreview
End Sub

Private Sub sltOutR_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    MapImageLevels sltInL.Value, sltInM.Value, sltInR.Value, sltOutL.Value, sltOutR.Value, True, fxPreview
End Sub
