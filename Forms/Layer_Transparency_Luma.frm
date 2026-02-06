VERSION 5.00
Begin VB.Form FormTransparency_FromLuma 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Generate mask from luminance"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11820
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   788
   Begin PhotoDemon.pdSlider sldGray 
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   3360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   873
      FontSizeCaption =   10
      Max             =   255
      Value           =   127
      NotchPosition   =   2
      NotchValueCustom=   127
   End
   Begin PhotoDemon.pdRadioButton rdoSource 
      Height          =   375
      Index           =   0
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "black is transparent"
      FontSize        =   12
      Value           =   -1  'True
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      ColorSelection  =   -1  'True
   End
   Begin PhotoDemon.pdRadioButton rdoSource 
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   2280
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "white is transparent"
      FontSize        =   12
   End
   Begin PhotoDemon.pdRadioButton rdoSource 
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   4
      Top             =   2880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "gray is transparent"
      FontSize        =   12
   End
End
Attribute VB_Name = "FormTransparency_FromLuma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Make layer transparent using luma data (dialog)
'Copyright 2020-2026 by Tanner Helland
'Created: 14/February/20
'Last updated: 14/February/20
'Last update: initial build
'
'This simple dialog will remap a layer's alpha values according to each pixel's luminance value.
' This is very helpful for overlaying textures atop a base layer; the texture's transparency can
' be auto-set according to luminance for a lot of nice effects - for example, load an image,
' create a duplicate layer, apply "Emboss" with neutral gray to the top layer, then use this tool
' with the "set gray to transparent" option - you get a very nice "relief" effect, helpful for
' sharpening (and with the added benefit of all layer blend modes + opacity for toggling).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK button
Private Sub cmdBar_OKClick()
    Process "Luminance to alpha", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdateDialogUI
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The user can also select a target color from the preview window
Private Sub pdFxPreview_ColorSelected()
    
    Dim targetColor As Long
    targetColor = pdFxPreview.SelectedColor
    
    'Set gray mode according to the selected color (and disable live previews - we'll reenable them
    ' manually at the end, after all changes are applied).
    cmdBar.SetPreviewStatus False
    If (targetColor = RGB(0, 0, 0)) Then
        rdoSource(0).Value = True
    ElseIf (targetColor = RGB(255, 255, 255)) Then
        rdoSource(1).Value = True
    Else
        rdoSource(2).Value = True
    End If
    
    sldGray.Value = Colors.GetLuminance(Colors.ExtractRed(targetColor), Colors.ExtractGreen(targetColor), Colors.ExtractBlue(targetColor))
    
    'Reenable live previews and immediately refresh the preview window
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

'Add transparency to an image by making a specified color transparent (chroma-key or "green screen").
' This function uses a high-quality color-matching scheme in the L*a*b* color space.
' LittleCMS is used for transforms, if present.
Public Sub LuminanceToAlpha(ByVal processParameters As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    
    Dim targetGray As Long
    
    With cParams
        targetGray = .GetLong("target-gray", vbBlack)
    End With
    
    If (Not toPreview) Then Message "Adding new alpha channel to image..."
    
    'Call prepImageData, which will prepare a temporary copy of the image
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'For this function to work, each pixel needs to be RGBA, 32-bpp
    Dim pxSize As Long
    pxSize = workingDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Build a lookup table for luminance values
    Dim newAlpha(0 To 255) As Byte
    For x = 0 To 255
    
        If (targetGray = 0) Then
            newAlpha(x) = x
        ElseIf (targetGray = 255) Then
            newAlpha(x) = 255 - x
        Else
            
            If (x < targetGray) Then
                newAlpha(x) = Int(CDbl((targetGray - x) / targetGray) * 255# + 0.5)
            Else
                newAlpha(x) = Int(CDbl((x - targetGray) / (255 - targetGray)) * 255# + 0.5)
            End If
            
        End If
    
    Next x
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'To improve performance of our horizontal loop, we'll move through bytes an entire pixel at a time
    Dim xStart As Long, xStop As Long
    xStart = initX * pxSize
    xStop = finalX * pxSize
        
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        
        'Wrap an array around the current scanline
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
        
        'With all lab values pre-calculated, we can quickly step through each pixel, calculating distances as we go
        For x = xStart To xStop Step pxSize
        
            'Get the source pixel color values
            b = imageData(x)
            g = imageData(x + 1)
            r = imageData(x + 2)
            
            'Calculate luminance
            targetGray = Colors.GetLuminance(r, g, b)
                
            'Assign the new alpha
            imageData(x + 3) = newAlpha(targetGray)
            
        Next x
        
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
        
    Next y
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

'Render a new preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then LuminanceToAlpha GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    Dim targetGray As Long
    If rdoSource(0).Value Then
        targetGray = 0
    ElseIf rdoSource(1).Value Then
        targetGray = 255
    Else
        targetGray = sldGray.Value
    End If
    
    With cParams
        .AddParam "target-gray", targetGray
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Sub rdoSource_Click(Index As Integer)
    UpdateDialogUI
    UpdatePreview
End Sub

Private Sub UpdateDialogUI()
    sldGray.Visible = rdoSource(2).Value
End Sub

Private Sub sldGray_Change()
    UpdatePreview
End Sub
