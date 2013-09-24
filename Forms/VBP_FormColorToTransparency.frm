VERSION 5.00
Begin VB.Form FormTransparency_FromColor 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Make color transparent"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11820
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
   ScaleWidth      =   788
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11820
      _ExtentX        =   20849
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
      ColorSelection  =   -1  'True
   End
   Begin PhotoDemon.sliderTextCombo sltErase 
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   2880
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   873
      Max             =   199
      Value           =   15
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
   Begin PhotoDemon.sliderTextCombo sltBlend 
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   3840
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   873
      Max             =   200
      Value           =   15
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
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   615
      Left            =   6240
      TabIndex        =   7
      Top             =   1800
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1085
      curColor        =   49152
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "edge blending:"
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
      TabIndex        =   6
      Top             =   3480
      Width           =   1590
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "erase threshold:"
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
      TabIndex        =   3
      Top             =   2520
      Width           =   1710
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "color to erase (click the preview to select):"
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
      TabIndex        =   2
      Top             =   1320
      Width           =   4485
   End
End
Attribute VB_Name = "FormTransparency_FromColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Make color transparent ("green screen") tool dialog
'Copyright ©2012-2013 by Tanner Helland
'Created: 13/August/13
'Last updated: 21/August/13
'Last update: gave this tool its own dialog, in keeping with the "do one thing and do it well" philosophy of PD dialogs
'
'PhotoDemon has long provided the ability to convert a 24bpp image to 32bpp, but the lack of an interface meant it could
' only add a fully opaque alpha channel.  Now the user can select from one of several conversion methods.
'
'This dialog present one of the more interesting conversion methods: a "color to alpha" technique, which allows for
' powerful green-screen capabilities.  A full CieLAB color space transformation is used, and an optional blend parameter
' will antialias and color-correct edges for maximum smoothness.  I don't know of any other software that utilizes this
' dual-threshold approach, and in my own testing, I have found PD to be superior to any other open source package at
' removing complex background colors.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Multiple controls on this form interact with each other; while they are interacting, disallow previews, then do a
' single preview after all controls have been set
Dim allowPreviews As Boolean

'OK button
Private Sub cmdBar_OKClick()
    Process "Color to alpha", , buildParams(colorPicker.Color, sltErase.Value, sltBlend.Value)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    colorPicker.Color = RGB(0, 192, 0)
    sltErase.Value = 15
    sltBlend.Value = 15
End Sub

Private Sub colorPicker_ColorChanged()
    updatePreview
End Sub

Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Render a preview of the alpha effect
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The user can select a color from the preview window; this helps green screen calculation immensely
Private Sub fxPreview_ColorSelected()
    colorPicker.Color = fxPreview.SelectedColor
    updatePreview
End Sub

'Convert a layer from 24bpp to 32bpp, based on the supplied convertType value:
' 0: use the supplied convertConstant value, and set the entire alpha channel to that
' 1: color-based.  Remove the color specified by convertColor, according to the thresholds supplied in eraseThreshold and blendThreshold
Public Sub colorToAlpha(Optional ByVal ConvertColor As Long, Optional ByVal eraseThreshold As Double = 15, Optional ByVal blendThreshold As Double = 30, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Adding new alpha channel to image..."
    
    'Call prepImageData, which will prepare a temporary copy of the image
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA, toPreview, dstPic
    
    'Before doing anything else, convert this layer to 32bpp.
    workingLayer.convertTo32bpp
    
    'Create a local array and point it at the pixel data we want to operate on
    prepSafeArray tmpSA, workingLayer
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
    qvDepth = workingLayer.getLayerColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'R2/G2/B2 store the RGB values of the color we are attempting to remove
    Dim r2 As Long, g2 As Long, b2 As Long
    r2 = ExtractR(ConvertColor)
    g2 = ExtractG(ConvertColor)
    b2 = ExtractB(ConvertColor)
    
    'For maximum quality, we will apply our color comparison in the CieLAB color space
    Dim labL As Double, labA As Double, labB As Double
    Dim labL2 As Double, labA2 As Double, labB2 As Double
    
    'Calculate the L*a*b* values of the color to be removed
    RGBtoLAB r2, g2, b2, labL2, labA2, labB2
    
    'The blend threshold is used to "smooth" the edges of the green screen.  Calculate the difference between
    ' the erase and the blend thresholds in advance.
    Dim difThreshold As Double
    blendThreshold = eraseThreshold + blendThreshold
    difThreshold = blendThreshold - eraseThreshold
    
    Dim cDistance As Double
    Dim newAlpha As Long
        
    'Loop through each pixel in the image, converting values as we go
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        'Convert the color to the L*a*b* color space
        RGBtoLAB r, g, b, labL, labA, labB
        
        'Perform a basic distance calculation (not ideal, but faster than a completely correct comparison;
        ' see http://en.wikipedia.org/wiki/Color_difference for a full report)
        cDistance = distanceThreeDimensions(labL, labA, labB, labL2, labA2, labB2)
        
        'If the distance is below the erasure threshold, remove it completely
        If cDistance < eraseThreshold Then
        
            ImageData(QuickVal + 3, Y) = 0
            
        'If the color is between the erasure and blend threshold, feather it against a partial alpha and
        ' color-correct it to remove any "color fringing" from the removed color.
        ElseIf cDistance < blendThreshold Then
            
            'Use a ^2 curve to improve blending response
            cDistance = ((blendThreshold - cDistance) / difThreshold)
            cDistance = cDistance * cDistance
            
            'Calculate a new alpha value for this pixel, based on its distance from the threshold.  Large
            ' distances from the removed color are made less transparent than small distances.
            newAlpha = 255 - (cDistance * 255)
            
            'Feathering the alpha often isn't enough to fully remove the color fringing caused by the removed
            ' background color, which will have "infected" the core RGB values.  Attempt to correct this by
            ' subtracting the target color from the original color, using the calculated threshold value; this
            ' is the only way I know to approximate the "feathering" caused by light bleeding over object edges.
            If cDistance = 1 Then cDistance = 0.999999
            r = (r - (r2 * cDistance)) / (1 - cDistance)
            g = (g - (g2 * cDistance)) / (1 - cDistance)
            b = (b - (b2 * cDistance)) / (1 - cDistance)
            
            If r > 255 Then r = 255
            If g > 255 Then g = 255
            If b > 255 Then b = 255
            If r < 0 Then r = 0
            If g < 0 Then g = 0
            If b < 0 Then b = 0
            
            'Assign the new color and alpha values
            ImageData(QuickVal + 2, Y) = r
            ImageData(QuickVal + 1, Y) = g
            ImageData(QuickVal, Y) = b
            ImageData(QuickVal + 3, Y) = newAlpha
                
        End If
        
    Next Y
        If Not toPreview Then
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
    
    'Enable/disable the remove/add alpha channel menus to match
    If Not toPreview Then metaToggle tImgMode32bpp, True

End Sub

Private Sub sltBlend_Change()
    updatePreview
End Sub

Private Sub sltErase_Change()
    updatePreview
End Sub

'Render a new preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then colorToAlpha colorPicker.Color, sltErase.Value, sltBlend.Value, True, fxPreview
End Sub
