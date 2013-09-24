VERSION 5.00
Begin VB.Form FormColorize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Colorize Options"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12345
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
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   823
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5820
      Width           =   12345
      _ExtentX        =   21775
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
   Begin PhotoDemon.smartCheckBox chkSaturation 
      Height          =   480
      Left            =   6240
      TabIndex        =   5
      Top             =   3240
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   847
      Caption         =   "preserve existing saturation"
      Value           =   1
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
   Begin VB.HScrollBar hsHue 
      Height          =   375
      Left            =   6240
      Max             =   359
      Min             =   1
      TabIndex        =   1
      Top             =   2160
      Value           =   180
      Width           =   5790
   End
   Begin VB.PictureBox picHueDemo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   2
      Top             =   2640
      Width           =   5310
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "color to apply:"
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
      Top             =   1680
      Width           =   1545
   End
End
Attribute VB_Name = "FormColorize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Colorize Form
'Copyright ©2006-2013 by Tanner Helland
'Created: 12/January/07
'Last updated: 09/September/12
'Last update: added "maintain saturation" check box
'
'Fairly simple and standard routine - look in the Miscellaneous Filters module
' for the HSL transformation code
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub cmdBar_OKClick()
    Process "Colorize", , buildParams(CDbl((CDbl(hsHue.Value) - 60) / 60), CBool(chkSaturation.Value))
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

'When the "maintain saturation" check box is clicked, redraw the image
Private Sub chkSaturation_Click()
    updatePreview
End Sub

'Colorize an image using a hue defined between -1 and 5
' Input: desired hue, whether to force saturation to 0.5 or maintain the existing value
Public Sub ColorizeImage(ByVal hToUse As Double, Optional ByVal maintainSaturation As Boolean = True, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Colorizing image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
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
    Dim h As Double, s As Double, l As Double
        
    'Loop through each pixel in the image, converting values as we go
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        'Get the hue and saturation
        tRGBToHSL r, g, b, h, s, l
        
        'Convert back to RGB using our artificial hue value
        If maintainSaturation Then
            tHSLToRGB hToUse, s, l, r, g, b
        Else
            tHSLToRGB hToUse, 0.5, l, r, g, b
        End If
        
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = g
        ImageData(QuickVal, Y) = b
        
    Next Y
        If toPreview = False Then
            If (X And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal X
            End If
        End If
    Next X
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reset the hue bar to the center position
Private Sub cmdBar_ResetClick()
    hsHue.Value = 180
End Sub

Private Sub Form_Activate()

    'This short routine is for drawing the picture box below the hue slider
    Dim hVal As Double
    Dim r As Long, g As Long, b As Long
    
    'Simple gradient-ish code implementation of drawing hue
    Dim X As Long
    For X = 0 To picHueDemo.ScaleWidth
    
        'Based on our x-position, gradient a value between -1 and 5
        hVal = X / picHueDemo.ScaleWidth
        hVal = hVal * 360
        hVal = (hVal - 60) / 60
        
        'Generate a hue for this position (the 1 and 0.5 correspond to full saturation and half luminance, respectively)
        tHSLToRGB hVal, 1, 0.5, r, g, b
        
        'Draw the color
        picHueDemo.Line (X, 0)-(X, picHueDemo.ScaleHeight), RGB(r, g, b)
        
    Next X
    
    picHueDemo.Picture = picHueDemo.Image
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When the hue scroll bar is changed, redraw the preview
Private Sub hsHue_Change()
    updatePreview
End Sub

Private Sub hsHue_Scroll()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ColorizeImage CSng((CSng(hsHue.Value) - 60) / 60), CBool(chkSaturation), True, fxPreview
End Sub
