VERSION 5.00
Begin VB.Form FormBlackLight 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Black light"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   Begin PhotoDemon.pdSlider sltIntensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2280
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "intensity"
      Min             =   1
      SigDigits       =   2
      Value           =   2
      DefaultValue    =   2
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
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
   End
End
Attribute VB_Name = "FormBlackLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Blacklight Form
'Copyright 2001-2026 by Tanner Helland
'Created: some time 2001
'Last updated: 01/October/13
'Last update: use a floating-point slider for more precise results
'
'I found this effect on accident, and it has gradually become one of my favorite effects.
' Visually stunning on many photographs.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Perform a blacklight filter
'Input: strength of the filter (min 1, no real max - but above 7 it becomes increasingly blown-out)
Public Sub fxBlackLight(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Illuminating image with imaginary blacklight..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim weight As Double
    weight = cParams.GetDouble("intensity", 2#)
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
    
    'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX
        
        'Get the source pixel color values
        xStride = x * 4
        b = imageData(xStride)
        g = imageData(xStride + 1)
        r = imageData(xStride + 2)
        
        'Calculate the gray value using the look-up table
        grayVal = grayLookUp(r + g + b)
        
        'Perform the blacklight conversion
        r = Abs(r - grayVal) * weight
        g = Abs(g - grayVal) * weight
        b = Abs(b - grayVal) * weight
        
        If (r > 255) Then r = 255
        If (g > 255) Then g = 255
        If (b > 255) Then b = 255
        
        'Assign that gray value to each color channel.  (Yes, channel order is reversed - that's deliberate!)
        imageData(xStride) = r
        imageData(xStride + 1) = g
        imageData(xStride + 2) = b
        
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
    Process "Black light", , GetLocalParamString, UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltIntensity.Value = 2#
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

'Update the preview whenever the combination slider/text control has its value changed
Private Sub sltIntensity_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then fxBlackLight GetLocalParamString, True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "intensity", sltIntensity.Value
    GetLocalParamString = cParams.GetParamString()
End Function
