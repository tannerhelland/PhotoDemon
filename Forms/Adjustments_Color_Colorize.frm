VERSION 5.00
Begin VB.Form FormColorize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Colorize"
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
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5820
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdCheckBox chkSaturation 
      Height          =   330
      Left            =   6240
      TabIndex        =   2
      Top             =   2760
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   582
      Caption         =   "preserve existing saturation"
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sltHue 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1800
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1270
      Caption         =   "color to apply"
      Max             =   359
      SliderTrackStyle=   4
      Value           =   180
      NotchPosition   =   1
      DefaultValue    =   180
   End
End
Attribute VB_Name = "FormColorize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Colorize Form
'Copyright 2006-2020 by Tanner Helland
'Created: 12/January/07
'Last updated: 22/June/14
'Last update: replace old scroll bar with slider/text combo
'
'Fairly simple and standard routine - look in the Miscellaneous Filters module
' for the HSL transformation code
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub cmdBar_OKClick()
    Process "Colorize", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

'When the "maintain saturation" check box is clicked, redraw the image
Private Sub chkSaturation_Click()
    UpdatePreview
End Sub

'Colorize an image using a hue defined between 0 and 359
' Input: desired hue, whether to force saturation to 0.5 or maintain the existing value
Public Sub ColorizeImage(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Colorizing image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim hToUse As Double, maintainSaturation As Boolean
    hToUse = cParams.GetDouble("hue", sltHue.Value)
    maintainSaturation = cParams.GetBool("preservesaturation", True)
    
    'Convert the incoming hue from [0, 360] to [-1, 5] range
    hToUse = hToUse / 60
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        xStride = x * 4
    For y = initY To finalY
    
        'Get the source pixel color values
        b = imageData(xStride, y)
        g = imageData(xStride + 1, y)
        r = imageData(xStride + 2, y)
        
        'Get the hue and saturation
        Colors.ImpreciseRGBtoHSL r, g, b, h, s, l
        
        'Convert back to RGB using our artificial hue value
        If maintainSaturation Then
            Colors.ImpreciseHSLtoRGB hToUse, s, l, r, g, b
        Else
            Colors.ImpreciseHSLtoRGB hToUse, 0.5, l, r, g, b
        End If
        
        'Assign the new values to each color channel
        imageData(xStride, y) = b
        imageData(xStride + 1, y) = g
        imageData(xStride + 2, y) = r
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

'Reset the hue bar to the center position
Private Sub cmdBar_ResetClick()
    sltHue.Value = 180#
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ColorizeImage GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltHue_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "hue", sltHue.Value
        .AddParam "preservesaturation", chkSaturation.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
