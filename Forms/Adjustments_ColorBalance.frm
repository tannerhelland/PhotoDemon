VERSION 5.00
Begin VB.Form FormColorBalance 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Adjust Color Balance"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12360
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
   ScaleWidth      =   824
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14802140
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltRed 
      Height          =   405
      Left            =   6000
      TabIndex        =   8
      Top             =   1800
      Width           =   6255
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   3
      GradientColorLeft=   16776960
      GradientColorRight=   255
      GradientColorMiddle=   8881021
   End
   Begin PhotoDemon.sliderTextCombo sltGreen 
      Height          =   405
      Left            =   6000
      TabIndex        =   9
      Top             =   2760
      Width           =   6255
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   3
      GradientColorLeft=   16711935
      GradientColorRight=   65280
   End
   Begin PhotoDemon.sliderTextCombo sltBlue 
      Height          =   405
      Left            =   6000
      TabIndex        =   10
      Top             =   3720
      Width           =   6255
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   3
      GradientColorLeft=   65535
      GradientColorRight=   16711680
   End
   Begin PhotoDemon.smartCheckBox chkLuminance 
      Height          =   360
      Left            =   6240
      TabIndex        =   12
      Top             =   4800
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   582
      Caption         =   "preserve luminance"
   End
   Begin PhotoDemon.buttonStrip btsTone 
      Height          =   600
      Left            =   6000
      TabIndex        =   14
      Top             =   540
      Width           =   6255
      _ExtentX        =   10425
      _ExtentY        =   1058
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "new balance"
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
      Left            =   5880
      TabIndex        =   13
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tonal range"
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
      Left            =   5880
      TabIndex        =   11
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "yellow"
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
      Height          =   270
      Left            =   6300
      TabIndex        =   7
      Top             =   4200
      Width           =   570
   End
   Begin VB.Label lblMagenta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "magenta"
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
      Height          =   270
      Left            =   6300
      TabIndex        =   6
      Top             =   3240
      Width           =   870
   End
   Begin VB.Label lblCyan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cyan"
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
      Height          =   270
      Left            =   6300
      TabIndex        =   5
      Top             =   2280
      Width           =   465
   End
   Begin VB.Label lblBlue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "blue"
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
      Height          =   270
      Left            =   10770
      TabIndex        =   3
      Top             =   4200
      Width           =   390
   End
   Begin VB.Label lblGreen 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "green"
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
      Height          =   270
      Left            =   10605
      TabIndex        =   2
      Top             =   3240
      Width           =   555
   End
   Begin VB.Label lblRed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "red"
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
      Height          =   270
      Left            =   10845
      TabIndex        =   1
      Top             =   2280
      Width           =   315
   End
End
Attribute VB_Name = "FormColorBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Balance Adjustment Form
'Copyright 2012-2015 by Tanner Helland & Audioglider
'Created: 31/January/13
'Last updated: 16/June/14
'Last update: Rewrote the color balance formula to allow the adjustment of
'             shadow/midtone/highlight tones.
'
'Fairly simple and standard color adjustment form.  Layout and feature set derived from comparable tools
' in GIMP and Photoshop.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Enum TONE_REGION
    TONE_SHADOWS = 0
    TONE_MIDTONES = 1
    TONE_HIGHLIGHTS = 2
End Enum

#If False Then
    Private Const TONE_SHADOWS = 0, TONE_MIDTONES = 1, TONE_HIGHLIGHTS = 2
#End If

'Apply a new color balance to the image
' Input: offset for each of red, green, and blue
Public Sub ApplyColorBalance(ByVal rVal As Long, ByVal gVal As Long, ByVal bVal As Long, ByVal nTone As Long, ByVal preserveLuminance As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Adjusting color balance..."
    
    Dim rModifier As Long, gModifier As Long, bModifier As Long
    rModifier = 0
    gModifier = 0
    bModifier = 0
    
    'Now, Build actual RGB modifiers based off the values provided
    gModifier = gModifier - rVal
    bModifier = bModifier - rVal
    rModifier = rModifier + rVal
    
    rModifier = rModifier - gVal
    bModifier = bModifier - gVal
    gModifier = gModifier + gVal
    
    rModifier = rModifier - bVal
    gModifier = gModifier - bVal
    bModifier = bModifier + bVal
    
    'Because these modifiers are constant throughout the image, we can build look-up tables for them
    Dim rLookup(0 To 255) As Byte, gLookUp(0 To 255) As Byte, bLookup(0 To 255) As Byte
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
    
    Dim rRgn(0 To 2) As Long, gRgn(0 To 2) As Long, bRgn(0 To 2) As Long
    
    rRgn(nTone) = rModifier
    gRgn(nTone) = gModifier
    bRgn(nTone) = bModifier
    
    'Add used for lightening, sub for darkening
    Dim highlightsAdd(0 To 255) As Double, midtonesAdd(0 To 255) As Double, shadowsAdd(0 To 255) As Double
    Dim highlightsSub(0 To 255) As Double, midtonesSub(0 To 255) As Double, shadowsSub(0 To 255) As Double
    
    Dim dl As Double, dm As Double
    
    For x = 0 To 255
        
        dl = 1.075 - 1 / (x / 16 + 1)
        
        dm = (x - 127) / 127
        dm = 0.667 * (1 - dm * dm)
        
        shadowsAdd(x) = dl
        shadowsSub(255 - x) = dl
        highlightsAdd(255 - x) = dl
        highlightsSub(x) = dl
        midtonesAdd(x) = dm
        midtonesSub(x) = dm
        
    Next x
    
    'Set up transfer arrays
    Dim rTransfer(0 To 2, 0 To 255) As Double, gTransfer(0 To 2, 0 To 255) As Double, bTransfer(0 To 2, 0 To 255) As Double
    
    'Add lighening/darkening modifiers to the transfer arrays
    For x = 0 To 255
        
        If rRgn(TONE_SHADOWS) > 0 Then rTransfer(TONE_SHADOWS, x) = shadowsAdd(x) Else rTransfer(TONE_SHADOWS, x) = shadowsSub(x)
        If rRgn(TONE_MIDTONES) > 0 Then rTransfer(TONE_MIDTONES, x) = midtonesAdd(x) Else rTransfer(TONE_MIDTONES, x) = midtonesSub(x)
        If rRgn(TONE_HIGHLIGHTS) > 0 Then rTransfer(TONE_HIGHLIGHTS, x) = highlightsAdd(x) Else rTransfer(TONE_HIGHLIGHTS, x) = highlightsSub(x)
    
        If gRgn(TONE_SHADOWS) > 0 Then gTransfer(TONE_SHADOWS, x) = shadowsAdd(x) Else gTransfer(TONE_SHADOWS, x) = shadowsSub(x)
        If gRgn(TONE_MIDTONES) > 0 Then gTransfer(TONE_MIDTONES, x) = midtonesAdd(x) Else gTransfer(TONE_MIDTONES, x) = midtonesSub(x)
        If gRgn(TONE_HIGHLIGHTS) > 0 Then gTransfer(TONE_HIGHLIGHTS, x) = highlightsAdd(x) Else gTransfer(TONE_HIGHLIGHTS, x) = highlightsSub(x)
    
        If bRgn(TONE_SHADOWS) > 0 Then bTransfer(TONE_SHADOWS, x) = shadowsAdd(x) Else bTransfer(TONE_SHADOWS, x) = shadowsSub(x)
        If bRgn(TONE_MIDTONES) > 0 Then bTransfer(TONE_MIDTONES, x) = midtonesAdd(x) Else bTransfer(TONE_MIDTONES, x) = midtonesSub(x)
        If bRgn(TONE_HIGHLIGHTS) > 0 Then bTransfer(TONE_HIGHLIGHTS, x) = highlightsAdd(x) Else bTransfer(TONE_HIGHLIGHTS, x) = highlightsSub(x)
    
    Next x
    
    'Populate the lookup tables
    For x = 0 To 255
        
        r = x
        g = x
        b = x
        
        'Apply the modifiers
        r = Clamp0255(r + (rRgn(TONE_SHADOWS) * rTransfer(TONE_SHADOWS, r)))
        r = Clamp0255(r + (rRgn(TONE_MIDTONES) * rTransfer(TONE_MIDTONES, r)))
        r = Clamp0255(r + (rRgn(TONE_HIGHLIGHTS) * rTransfer(TONE_HIGHLIGHTS, r)))
        
        g = Clamp0255(g + (gRgn(TONE_SHADOWS) * gTransfer(TONE_SHADOWS, g)))
        g = Clamp0255(g + (gRgn(TONE_MIDTONES) * gTransfer(TONE_MIDTONES, g)))
        g = Clamp0255(g + (gRgn(TONE_HIGHLIGHTS) * gTransfer(TONE_HIGHLIGHTS, g)))
        
        b = Clamp0255(b + (bRgn(TONE_SHADOWS) * bTransfer(TONE_SHADOWS, b)))
        b = Clamp0255(b + (bRgn(TONE_MIDTONES) * bTransfer(TONE_MIDTONES, b)))
        b = Clamp0255(b + (bRgn(TONE_HIGHLIGHTS) * bTransfer(TONE_HIGHLIGHTS, b)))
        
        rLookup(x) = r
        gLookUp(x) = g
        bLookup(x) = b
    
    Next x
    
    Dim origLuminance As Double
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Get the original luminance
        origLuminance = getLuminance(r, g, b) / 255
        
        r = rLookup(r)
        g = gLookUp(g)
        b = bLookup(b)
        
        'If the user doesn't want us to maintain luminance, our work is done - assign the new values.
        'If they do want us to maintain luminance, things are a bit trickier.  We need to convert our values to
        ' HSL, then substitute the original luminance and convert back to RGB.
        If preserveLuminance Then
        
            'Convert the new values to HSL
            tRGBToHSL r, g, b, h, s, l
            
            'Now, convert back, using the original luminance
            tHSLToRGB h, s, origLuminance, r, g, b
            
        End If
        
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Limit color to a 0-255 range
Private Function Clamp0255(ByVal d As Double) As Double
    If d < 255 Then
        If d > 0 Then Clamp0255 = d Else Clamp0255 = 0
        Exit Function
    End If
    Clamp0255 = 255
End Function

Private Sub btsTone_Click(ByVal buttonIndex As Long)
    updatePreview
End Sub

Private Sub chkLuminance_Click()
    updatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Color balance", , buildParams(sltRed, sltGreen, sltBlue, btsTone.ListIndex, CBool(chkLuminance.Value)), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRed.Value = 0
    sltGreen.Value = 0
    sltBlue.Value = 0
    btsTone.ListIndex = 1 'Default to midtone correction
    chkLuminance.Value = vbChecked
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    makeFormPretty Me
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Populate the button strip
    btsTone.AddItem "shadows", 0
    btsTone.AddItem "midtones", 1
    btsTone.AddItem "highlights", 2
    btsTone.ListIndex = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltBlue_Change()
    updatePreview
End Sub

Private Sub sltGreen_Change()
    updatePreview
End Sub

Private Sub sltRed_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyColorBalance sltRed, sltGreen, sltBlue, btsTone.ListIndex, CBool(chkLuminance), True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub
