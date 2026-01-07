VERSION 5.00
Begin VB.Form FormSepia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Sepia"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
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
   ScaleWidth      =   768
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
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sldStrength 
      CausesValidation=   0   'False
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2400
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "strength"
      Max             =   100
      GradientColorLeft=   15102446
      GradientColorRight=   8253041
      GradientColorMiddle=   16777215
      NotchPosition   =   2
      NotchValueCustom=   100
   End
End
Attribute VB_Name = "FormSepia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Sepia Dialog
'Copyright 2014-2026 by Tanner Helland
'Created: 03/July/14
'Last updated: 27/April/20
'Last update: add dialog so the user can adjust "strength" (instead of being locked at 100%)
'
'Sepia uses the W3C formula for tinting; values derived from the standard at:
' https://dvcs.w3.org/hg/FXTF/raw-file/tip/filters/index.html#sepiaEquivalent
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub ApplySepiaEffect(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Engaging hipsters to perform sepia conversion..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim strength As Single
    strength = cParams.GetSingle("strength", 100!) / 100!
    
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
    Dim newR As Double, newG As Double, newB As Double
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Apply the sepia transform.  Note that we manually applying nearest-int rounding
        newR = Int(r * 0.393 + 0.5) + Int(g * 0.769 + 0.5) + Int(b * 0.189 + 0.5)
        newG = Int(r * 0.349 + 0.5) + Int(g * 0.686 + 0.5) + Int(b * 0.168 + 0.5)
        newB = Int(r * 0.272 + 0.5) + Int(g * 0.534 + 0.5) + Int(b * 0.131 + 0.5)
        
        'The w3c transform can overflow; correct immediately or the strength blend can fail
        If (newB > 255) Then newB = 255
        If (newG > 255) Then newG = 255
        If (newR > 255) Then newR = 255
        
        'Apply strength
        If (strength < 1!) Then
            newB = Colors.BlendColors(b, newB, strength)
            newG = Colors.BlendColors(g, newG, strength)
            newR = Colors.BlendColors(r, newR, strength)
        End If
        
        'Assign new values
        imageData(x) = newB
        imageData(x + 1) = newG
        imageData(x + 2) = newR
        
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
    Process "Sepia", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sldStrength.Value = 100
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

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplySepiaEffect GetLocalParamString(), True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "strength", sldStrength.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Sub sldStrength_Change()
    UpdatePreview
End Sub
