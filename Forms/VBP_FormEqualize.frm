VERSION 5.00
Begin VB.Form FormEqualize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Equalize Histogram"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10155
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
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   677
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5805
      Width           =   10155
      _ExtentX        =   17912
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
   Begin PhotoDemon.smartCheckBox chkRed 
      Height          =   570
      Left            =   6240
      TabIndex        =   3
      Top             =   2040
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1005
      Caption         =   "red"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartCheckBox chkGreen 
      Height          =   570
      Left            =   6240
      TabIndex        =   4
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1005
      Caption         =   "green"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartCheckBox chkBlue 
      Height          =   570
      Left            =   6240
      TabIndex        =   5
      Top             =   3000
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1005
      Caption         =   "blue"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartCheckBox chkLuminance 
      Height          =   570
      Left            =   6240
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1005
      Caption         =   "luminance"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblEqualize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "equalize:"
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
      TabIndex        =   1
      Top             =   1620
      Width           =   945
   End
End
Attribute VB_Name = "FormEqualize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Histogram Equalization Interface
'Copyright ©2012-2013 by Tanner Helland
'Created: 19/September/12
'Last updated: 22/August/13
'Last update: add command bar user control
'
'Module for handling histogram equalization.  Any combination of red, green, blue, and luminance can be equalized, but if
' luminance is selected it will get precedent (e.g. it will be equalized first).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Whenever a check box is changed, redraw the preview
Private Sub chkBlue_Click()
    updatePreview
End Sub

Private Sub chkGreen_Click()
    updatePreview
End Sub

Private Sub chkLuminance_Click()
    updatePreview
End Sub

Private Sub chkRed_Click()
    updatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Equalize", , buildParams(CBool(chkRed), CBool(chkGreen), CBool(chkBlue), CBool(chkLuminance))
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Request a preview
    updatePreview
    
End Sub

'Equalize the red, green, blue, and/or Luminance channels of an image
' (Technically Luminance isn't a channel, but you know what I mean.)
Public Sub EqualizeHistogram(ByVal HandleR As Boolean, ByVal HandleG As Boolean, ByVal HandleB As Boolean, ByVal HandleL As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Analyzing image histogram..."
    
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
    SetProgBarMax finalX * 2
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
    Dim lInt As Long
    
    'Histogram variables
    Dim rData(0 To 255) As Double, gData(0 To 255) As Double, bData(0 To 255) As Double
    Dim rDataInt(0 To 255) As Long, gDataInt(0 To 255) As Long, bDataInt(0 To 255) As Long
    Dim lData(0 To 255) As Double
    Dim lDataInt(0 To 255) As Long
        
    'Loop through each pixel in the image, converting values as we go.
    ' (This step is so fast that I calculate all channels, even those not being converted, with the exception of luminance.)
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        'Store those values in the histogram
        rDataInt(r) = rDataInt(r) + 1
        gDataInt(g) = gDataInt(g) + 1
        bDataInt(b) = bDataInt(b) + 1
        
        'Because luminance is slower to calculate, only calculate it if absolutely necessary
        If HandleL Then
            lInt = getLuminance(r, g, b)
            lDataInt(lInt) = lDataInt(lInt) + 1
        End If
        
    Next Y
        If toPreview = False Then
            If (X And progBarCheck) = 0 Then SetProgBarVal X
        End If
    Next X
    
    'Compute a scaling factor based on the number of pixels in the image
    Dim scaleFactor As Double
    scaleFactor = 255 / (curLayerValues.Width * curLayerValues.Height)
    
    'Compute red if requested
    If HandleR Then
        rData(0) = rDataInt(0) * scaleFactor
        For X = 1 To 255
            rData(X) = rData(X - 1) + (scaleFactor * rDataInt(X))
        Next X
    End If
    
    'Compute green if requested
    If HandleG Then
        gData(0) = gDataInt(0) * scaleFactor
        For X = 1 To 255
            gData(X) = gData(X - 1) + (scaleFactor * gDataInt(X))
        Next X
    End If
    
    'Compute blue if requested
    If HandleB Then
        bData(0) = bDataInt(0) * scaleFactor
        For X = 1 To 255
            bData(X) = bData(X - 1) + (scaleFactor * bDataInt(X))
        Next X
    End If
    
    'Compute luminance if requested
    If HandleL Then
        lData(0) = lDataInt(0) * scaleFactor
        For X = 1 To 255
            lData(X) = lData(X - 1) + (scaleFactor * lDataInt(X))
        Next X
    End If
    
    'Make sure all look-up values are in valid byte range (e.g. [0,255])
    For X = 0 To 255
        
        If rData(X) > 255 Then
            rDataInt(X) = 255
        Else
            rDataInt(X) = Int(rData(X))
        End If
        
        If gData(X) > 255 Then
            gDataInt(X) = 255
        Else
            gDataInt(X) = Int(gData(X))
        End If
        
        If bData(X) > 255 Then
            bDataInt(X) = 255
        Else
            bDataInt(X) = Int(bData(X))
        End If
        
        If lData(X) > 255 Then
            lDataInt(X) = 255
        Else
            lDataInt(X) = Int(lData(X))
        End If
        
    Next X
    
    'Apply the equalized values
    If Not toPreview Then Message "Equalizing image..."
    
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
    
        'Get the RGB values
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        'If luminance has been requested, calculate it before messing with any of the color channels
        If HandleL Then
            tRGBToHSL r, g, b, h, s, l
            tHSLToRGB h, s, lDataInt(Int(l * 255)) / 255, r, g, b
        End If
        
        'Next, calculate new values for the color channels, based on what is being equalized
        If HandleR Then r = rDataInt(r)
        If HandleG Then g = gDataInt(g)
        If HandleB Then b = bDataInt(b)
        
        'Assign our new values back into the pixel array
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = g
        ImageData(QuickVal, Y) = b
        
    Next Y
        If toPreview = False Then
            If (X And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal X + finalX
            End If
        End If
    Next X
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then EqualizeHistogram CBool(chkRed.Value), CBool(chkGreen.Value), CBool(chkBlue.Value), CBool(chkLuminance.Value), True, fxPreview
End Sub
