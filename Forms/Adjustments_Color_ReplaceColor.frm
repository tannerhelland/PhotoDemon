VERSION 5.00
Begin VB.Form FormReplaceColor 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Replace color"
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   1323
      BackColor       =   14802140
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
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   3000
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1270
      Caption         =   "replace threshold"
      Max             =   199
      Value           =   15
   End
   Begin PhotoDemon.sliderTextCombo sltBlend 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   4080
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1270
      Caption         =   "edge blending"
      Max             =   200
      Value           =   15
   End
   Begin PhotoDemon.colorSelector colorOld 
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      Caption         =   "color to replace (right-click preview to select)"
      curColor        =   12582912
   End
   Begin PhotoDemon.colorSelector colorNew 
      Height          =   975
      Left            =   6000
      TabIndex        =   5
      Top             =   1680
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      Caption         =   "new color"
      curColor        =   49152
   End
End
Attribute VB_Name = "FormReplaceColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Replace color dialog
'Copyright 2013-2015 by Tanner Helland
'Created: 29/October/13
'Last updated: 30/October/13
'Last update: finished initial build
'
'This function uses an algorithm very similar to PhotoDemon's green screen (FormTransparency_FromColor) algorithm.
' Separate sliders are provided for both a replacement threshold, and a blend threshold, to help the user minimize
' harsh edges between the color and its surroundings.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'OK button
Private Sub cmdBar_OKClick()
    Process "Replace color", , buildParams(colorOld.Color, colorNew.Color, sltErase.Value, sltBlend.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    colorOld.Color = RGB(0, 0, 255)
    colorNew.Color = RGB(0, 192, 0)
    sltErase.Value = 15
    sltBlend.Value = 15
End Sub

Private Sub colorNew_ColorChanged()
    updatePreview
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Render a preview of the alpha effect
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The user can select a color from the preview window; this helps green screen calculation immensely
Private Sub fxPreview_ColorSelected()
    colorOld.Color = fxPreview.SelectedColor
    updatePreview
End Sub

'Replace one color in an image with another color, with full blending and feathering support
Public Sub ReplaceSelectedColor(ByVal oldColor As Long, ByVal newColor As Long, Optional ByVal eraseThreshold As Double = 15, Optional ByVal blendThreshold As Double = 30, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Replacing color..."
    
    'Call prepImageData, which will prepare a temporary copy of the image
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA, toPreview, dstPic
        
    'Create a local array and point it at the pixel data we want to operate on
    prepSafeArray tmpSA, workingDIB
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
    qvDepth = workingDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'oldR/G/B store the RGB values of the color we are attempting to remove
    Dim oldR As Long, oldG As Long, oldB As Long
    oldR = ExtractR(oldColor)
    oldG = ExtractG(oldColor)
    oldB = ExtractB(oldColor)
    
    'newR/G/B store the RGB values of the color we are using to replace the old color
    Dim newR As Long, newG As Long, newB As Long
    newR = ExtractR(newColor)
    newG = ExtractG(newColor)
    newB = ExtractB(newColor)
    
    'For maximum quality, we will apply our color comparison in the CieLAB color space
    Dim labL As Double, labA As Double, labB As Double
    Dim labL2 As Double, labA2 As Double, labB2 As Double
    
    'Calculate the L*a*b* values of the color to be replaced
    RGBtoLAB oldR, oldG, oldB, labL2, labA2, labB2
    
    'The blend threshold is used to "smooth" the edges of replaced color areas.  Calculate the difference between
    ' the erase and the blend thresholds in advance.
    Dim difThreshold As Double
    blendThreshold = eraseThreshold + blendThreshold
    difThreshold = blendThreshold - eraseThreshold
    
    Dim cDistance As Double
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Convert the color to the L*a*b* color space
        RGBtoLAB r, g, b, labL, labA, labB
        
        'Perform a basic distance calculation (not ideal, but faster than a completely correct comparison;
        ' see http://en.wikipedia.org/wiki/Color_difference for a full report)
        cDistance = distanceThreeDimensions(labL, labA, labB, labL2, labA2, labB2)
        
        'If the distance is below the erasure threshold, replace it completely
        If cDistance < eraseThreshold Then
        
            ImageData(QuickVal + 2, y) = newR
            ImageData(QuickVal + 1, y) = newG
            ImageData(QuickVal, y) = newB
            
        'If the color is between the replace and blend threshold, feather it against the new color and
        ' color-correct it to remove any "color fringing" from the replaced color.
        ElseIf cDistance < blendThreshold Then
            
            'Use a ^2 curve to improve blending response
            cDistance = ((blendThreshold - cDistance) / difThreshold)
            
            'Feathering the pixel often isn't enough to fully remove the color fringing caused by the replaced
            ' color, which will have "infected" the core RGB values.  Attempt to correct this by subtracting the
            ' target color from the original color, using the calculated threshold value; this is the only way I
            ' know to approximate the "feathering" caused by light bleeding over object edges.
            If cDistance = 1 Then cDistance = 0.999999
            r = (r - (oldR * cDistance)) / (1 - cDistance)
            g = (g - (oldG * cDistance)) / (1 - cDistance)
            b = (b - (oldB * cDistance)) / (1 - cDistance)
            
            If r > 255 Then r = 255
            If g > 255 Then g = 255
            If b > 255 Then b = 255
            If r < 0 Then r = 0
            If g < 0 Then g = 0
            If b < 0 Then b = 0
            
            'Assign the new color and alpha values
            ImageData(QuickVal + 2, y) = BlendColors(r, newR, cDistance)
            ImageData(QuickVal + 1, y) = BlendColors(g, newG, cDistance)
            ImageData(QuickVal, y) = BlendColors(b, newB, cDistance)
                
        End If
        
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

Private Sub sltBlend_Change()
    updatePreview
End Sub

Private Sub sltErase_Change()
    updatePreview
End Sub

'Render a new preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ReplaceSelectedColor colorOld.Color, colorNew.Color, sltErase.Value, sltBlend.Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

