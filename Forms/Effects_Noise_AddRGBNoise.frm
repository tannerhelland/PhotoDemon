VERSION 5.00
Begin VB.Form FormNoise 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Noise"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11655
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
   ScaleWidth      =   777
   Begin PhotoDemon.pdButtonStrip btsColor 
      Height          =   1095
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   1931
      Caption         =   "appearance"
      FontSize        =   11
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11655
      _ExtentX        =   20558
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
   End
   Begin PhotoDemon.pdSlider sltNoise 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "strength"
      Max             =   100
      SigDigits       =   1
      Value           =   10
      DefaultValue    =   10
   End
   Begin PhotoDemon.pdButtonStrip btsDistribution 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   3240
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   1931
      Caption         =   "distribution"
      FontSize        =   11
   End
End
Attribute VB_Name = "FormNoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Noise Interface
'Copyright 2001-2026 by Tanner Helland
'Created: 15/March/01
'Last updated: 07/August/17
'Last update: large performance and quality improvements, Gaussian noise option, convert to XML params
'
'Want to add artifical noise to an image?  If so, you've come to the right place.  This dialog allows the user
' to add various types of noise to an image, including monochrome or color noise, in uniform or normal distributions.
' THe pdRandomize class does most the heavy lifting, and this dialog could easily be overhauled to expose a seed
' value to the user.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub AddNoise(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Increasing image noise..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim noiseAmount As Double, useMono As Boolean, useGaussian As Boolean
    
    With cParams
        noiseAmount = .GetDouble("amount", sltNoise.Value)
        useMono = .GetBool("monochrome", False)
        useGaussian = .GetBool("gaussian", False)
    End With
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right * 4
    finalY = curDIBValues.Bottom
    
    Dim dibPtr As Long, dibStride As Long
    dibPtr = workingDIB.GetDIBPointer
    dibStride = workingDIB.GetDIBStride
    workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, 0
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim nColor As Long
    
    'noiseAmount is returned on the range [0.0, 100.0].  At maximum strength, we want to scale this to
    ' [-255.0, 255.0], or a large enough number to turn white pixels black (and vice-versa).
    noiseAmount = noiseAmount * 2.55
    
    If useGaussian Then
        noiseAmount = noiseAmount * 0.333
    Else
        noiseAmount = noiseAmount * 2#
    End If
    
    '(Note that the rest of the scaling takes place inside the actual loop, as the random noise generator only
    ' produces positive numbers, obviously.)
    Dim noiseOffset As Double
    If (Not useGaussian) Then noiseOffset = noiseAmount * -0.5
    
    'Although it's slow, we're stuck using random numbers for noise addition.  Seed the generator with a pseudo-random value.
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_AutomaticAndRandom
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        tmpSA1D.pvData = dibPtr + y * dibStride
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Monochromatic noise - same amount for each color
        If useMono Then
            
            If useGaussian Then
                nColor = noiseAmount * cRandom.GetGaussianFloat_WH()
            Else
                nColor = noiseOffset + noiseAmount * cRandom.GetRandomFloat_WH()
            End If
            
            r = r + nColor
            g = g + nColor
            b = b + nColor
        
        'Colored noise - each color generated randomly
        Else
            
            If useGaussian Then
                r = r + noiseAmount * cRandom.GetGaussianFloat_WH()
                g = g + noiseAmount * cRandom.GetGaussianFloat_WH()
                b = b + noiseAmount * cRandom.GetGaussianFloat_WH()
            Else
                r = r + noiseOffset + noiseAmount * cRandom.GetRandomFloat_WH()
                g = g + noiseOffset + noiseAmount * cRandom.GetRandomFloat_WH()
                b = b + noiseOffset + noiseAmount * cRandom.GetRandomFloat_WH()
            End If
            
        End If
        
        If (r > 255) Then r = 255
        If (r < 0) Then r = 0
        If (g > 255) Then g = 255
        If (g < 0) Then g = 0
        If (b > 255) Then b = 255
        If (b < 0) Then b = 0
        
        'Assign that blended value to each color channel
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

Private Sub btsColor_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsDistribution_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Add RGB noise", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Populate coloring options
    btsColor.AddItem "color", 0
    btsColor.AddItem "monochrome", 1
    btsColor.ListIndex = 0
    
    btsDistribution.AddItem "uniform", 0
    btsDistribution.AddItem "gaussian", 1
    btsDistribution.ListIndex = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltNoise_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.AddNoise GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "amount", sltNoise.Value
        .AddParam "monochrome", (btsColor.ListIndex = 1)
        .AddParam "gaussian", (btsDistribution.ListIndex = 1)
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
