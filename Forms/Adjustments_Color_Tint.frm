VERSION 5.00
Begin VB.Form FormTint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Tint"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11280
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
   ScaleWidth      =   752
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltTint 
      CausesValidation=   0   'False
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2400
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1244
      Caption         =   "tint"
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   3
      GradientColorLeft=   15102446
      GradientColorRight=   8253041
      GradientColorMiddle=   16777215
   End
End
Attribute VB_Name = "FormTint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tint Dialog
'Copyright 2014-2026 by Tanner Helland
'Created: 03/July/14
'Last updated: 20/July/17
'Last update: migrate to XML params
'
'Tint is a simple adjustment along the magenta/green axis of the image.  While of limited use in a
' separate dialog like this, PhotoDemon sticks to convention by providing it as a "quick-fix" non-destructive
' action, which also means that it needs to exist as a dedicated menu entry.
'
'The formula used here is more nuanced than the "quick fix" version.  This tool will attempt to preserve image
' luminance, by compensating for the loss (or gain) of green light via adjustments to the red and blue channels.
' This provides a better end result, but note that it *will* differ from a matching adjustment via the
' tint quick-fix slider.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Change the tint of an image
' INPUT: tint adjustment, [-100, 100], 0 = no change
Public Sub AdjustTint(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Re-tinting image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim tintAdjustment As Long
    tintAdjustment = cParams.GetDouble("tint", 0#)
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left * 4
    initY = curDIBValues.Top
    finalX = curDIBValues.Right * 4
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, v As Double, origV As Double
    
    'Build a look-up table of tint values.  (Tint only affects the green channel)
    Dim gLookup() As Long
    ReDim gLookup(0 To 255) As Long
    For x = 0 To 255
        g = x + (tintAdjustment \ 2)
        If (g > 255) Then g = 255
        If (g < 0) Then g = 0
        gLookup(x) = g
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Calculate luminance
        origV = Colors.GetLuminance(r, g, b) / 255#
        
        'Convert the re-tinted colors to HSL
        Colors.ImpreciseRGBtoHSL r, gLookup(g), b, h, s, v
        
        'Convert back to RGB
        Colors.ImpreciseHSLtoRGB h, s, origV, r, g, b
        
        'Assign new values
        imageData(x) = b
        imageData(x + 1) = g
        imageData(x + 2) = r
        
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

Private Sub cmdBar_OKClick()
    Process "Tint", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltTint_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.AdjustTint GetLocalParamString(), True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "tint", sltTint.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
