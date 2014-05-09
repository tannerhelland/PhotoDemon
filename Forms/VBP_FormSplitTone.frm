VERSION 5.00
Begin VB.Form FormSplitTone 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Split Toning"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   285
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
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   5
      Top             =   5730
      Width           =   12090
      _ExtentX        =   21325
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
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltBalance 
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   2295
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
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
   Begin PhotoDemon.colorSelector cpHighlight 
      Height          =   615
      Left            =   6240
      TabIndex        =   6
      Top             =   1080
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1085
      curColor        =   16744192
   End
   Begin PhotoDemon.colorSelector cpShadow 
      Height          =   615
      Left            =   6240
      TabIndex        =   7
      Top             =   3480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1085
      curColor        =   32767
   End
   Begin PhotoDemon.sliderTextCombo sltStrength 
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   4665
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      Max             =   100
      Value           =   100
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
   Begin VB.Label lblStrength 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "toning strength:"
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
      TabIndex        =   9
      Top             =   4320
      Width           =   1710
   End
   Begin VB.Label lblShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "shadow color:"
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
      TabIndex        =   4
      Top             =   3060
      Width           =   1500
   End
   Begin VB.Label lblHue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "highlight color:"
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
      Top             =   690
      Width           =   1620
   End
   Begin VB.Label lblBalance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "balance:"
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
      TabIndex        =   0
      Top             =   1950
      Width           =   885
   End
End
Attribute VB_Name = "FormSplitTone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Split Toning Form
'Copyright ©2014 by Audioglider and Tanner Helland
'Created: 07/May/14
'Last updated: 09/May/14
'Last update: switch to HSL instead of HSV; this is slower, but arguably more appropriate for this tool.
'
'This technique applies a different tones to shadows and highlights in the image.  For a comprehensive explanation
' of split-toning (and its historical relevance), see
' http://www.alternativephotography.com/wp/toning/split-toning-history
'
'Many thanks to expert coder Audioglider for contributing this tool to PhotoDemon.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Apply a split-tone filter to the current layer or selection
'Inputs:
'  - Highlight color (as Long, created via VB's RGB() command)
'  - Shadow color (as Long, created via VB's RGB() command)
'  - Balance parameter, [-100, 100].  At 0, tones will be equally split between the highlight and shadow colors.  > 0 Balance will favor
'     highlights, while < 0 will favor shadows.
'  - Strength parameter, [0, 100].  At 100, current pixel values will be overwritten by their split-toned counterparts.  At 50, the original
'     and split-toned RGB values will be blended at a 50/50 ratio.  0 = no change.
Public Sub SplitTone(ByVal highlightColor As Long, ByVal shadowColor As Long, ByVal Balance As Double, ByVal Strength As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Split-toning image..."
    
    'From the incoming colors, determine corresponding hue and saturation values
    Dim highlightHue As Double, highlightSaturation As Double, shadowHue As Double, shadowSaturation As Double
    Dim ignoreLuminance As Double
    fRGBtoHSL ExtractR(highlightColor) / 255, ExtractG(highlightColor) / 255, ExtractB(highlightColor) / 255, highlightHue, highlightSaturation, ignoreLuminance
    fRGBtoHSL ExtractR(shadowColor) / 255, ExtractG(shadowColor) / 255, ExtractB(shadowColor) / 255, shadowHue, shadowSaturation, ignoreLuminance
    
    'Convert balance mix value to [0,2]; it will be used to blend split-toned colors at a varying scale (low balance
    ' favors the shadow tone, high balance favors the highlight tone)
    Dim balGradient As Double
    balGradient = Math_Functions.convertRange(-100, 100, 0, 2, Balance)
    
    'Strength controls the ratio at which the split-toned pixels are merged with the original pixels.  We want it on a [0, 1] scale.
    Strength = Math_Functions.convertRange(0, 100, 0, 1, Strength)
    
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
    Dim newR As Long, newG As Long, newB As Long
    Dim v As Double
    
    Dim rHighlight As Double, gHighlight As Double, bHighlight As Double
    Dim rShadow As Double, gShadow As Double, bShadow As Double
    Dim thisGradient As Double
    
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate HSL-compatible luminance
        v = getLuminance(r, g, b) / 255
        
        'Retrieve RGB conversions for the supplied highlight and shadow values, but retaining the pixel's current luminance (v)
        fHSLtoRGB highlightHue, highlightSaturation, v, rHighlight, gHighlight, bHighlight
        fHSLtoRGB shadowHue, shadowSaturation, v, rShadow, gShadow, bShadow
        
        'Highlight and shadow values are returned in the range [0, 1]; convert them to [0, 255] before continuing
        rHighlight = rHighlight * 255
        rShadow = rShadow * 255
        gHighlight = gHighlight * 255
        gShadow = gShadow * 255
        bHighlight = bHighlight * 255
        bShadow = bShadow * 255
        
        'We now have shadow and highlight colors for this pixel, already modified according to this pixel's luminance.
        
        'Next, we need to decide the ratio at which to mix the colors.  This is controlled by the balance slider; at a Balance of 0,
        ' the colors are equally mixed between the shadow and gradient colors according to their luminance.  If the Balance is > 0,
        ' we favor the highlight color, and if < 0 we favor the shadow color.
        thisGradient = v * balGradient
        If thisGradient > 1 Then thisGradient = 1
        If thisGradient < 0 Then thisGradient = 0
        
        'Use the balance value to mix the shadow and highlight colors
        newR = BlendColors(rShadow, rHighlight, thisGradient)
        newG = BlendColors(gShadow, gHighlight, thisGradient)
        newB = BlendColors(bShadow, bHighlight, thisGradient)
        
        'Finally, apply the new RGB values to the image by blending them with their original color at the user's requested strength.
        ImageData(QuickVal + 2, y) = BlendColors(r, newR, Strength)
        ImageData(QuickVal + 1, y) = BlendColors(g, newG, Strength)
        ImageData(QuickVal, y) = BlendColors(b, newB, Strength)
        
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

Private Sub cmdBar_OKClick()
    Process "Split toning", , buildParams(cpHighlight.Color, cpShadow.Color, sltBalance.Value, sltStrength.Value)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

'To help orient the user, slightly different reset values are used for this tool.
Private Sub cmdBar_ResetClick()
    cpHighlight.Color = RGB(150, 200, 255)
    cpShadow.Color = RGB(255, 200, 150)
    sltBalance.Value = 0
    sltStrength.Value = 100
End Sub

Private Sub cpHighlight_ColorChanged()
    updatePreview
End Sub

Private Sub cpShadow_ColorChanged()
    updatePreview
End Sub

Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then SplitTone cpHighlight.Color, cpShadow.Color, sltBalance.Value, sltStrength.Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltStrength_Change()
    updatePreview
End Sub

Private Sub sltBalance_Change()
    updatePreview
End Sub
