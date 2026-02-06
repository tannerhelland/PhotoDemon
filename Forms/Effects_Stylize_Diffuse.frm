VERSION 5.00
Begin VB.Form FormDiffuse 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Diffuse"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11790
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
   ScaleWidth      =   786
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltX 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1244
      Caption         =   "horizontal strength"
      Max             =   100
      SigDigits       =   2
      Value           =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdCheckBox chkWrap 
      Height          =   330
      Left            =   6120
      TabIndex        =   2
      Top             =   3600
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   582
      Caption         =   "wrap edge values"
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
   Begin PhotoDemon.pdSlider sltY 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   2640
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1244
      Caption         =   "vertical strength"
      Max             =   100
      SigDigits       =   2
      Value           =   1
      DefaultValue    =   1
   End
End
Attribute VB_Name = "FormDiffuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Diffuse Filter Handler
'Copyright 2001-2026 by Tanner Helland
'Created: 8/14/01
'Last updated: 08/August/17
'Last update: migrate to XML params, large performance improvements
'
'Module for handling "diffuse"-style filters (also called "displace", e.g. in GIMP).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub chkWrap_Click()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Diffuse", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews until everything is loaded
    cmdBar.SetPreviewStatus False
     
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Custom diffuse effect
' Inputs: diameter in x direction, diameter in y direction, whether or not to wrap edge pixels, and optional preview settings
Public Sub DiffuseCustom(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Simulating large image explosion..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim xDiffuse As Long, yDiffuse As Long, wrapPixels As Boolean
    Dim xDiffuseRatio As Double, yDiffuseRatio As Double
    
    With cParams
        xDiffuseRatio = .GetDouble("xsize", sltX.Value)
        yDiffuseRatio = .GetDouble("ysize", sltY.Value)
        wrapPixels = .GetBool("wrap", False)
    End With
    
    'Remap the diffuse ratios to the scale [0, 1]
    xDiffuseRatio = 0.01 * xDiffuseRatio
    yDiffuseRatio = 0.01 * yDiffuseRatio
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Long, dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, , , True
    workingDIB.WrapLongArrayAroundDIB dstImageData, dstSA
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    Dim srcImageData() As Long, srcSA As SafeArray2D
    srcDIB.WrapLongArrayAroundDIB srcImageData, srcSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'Scale the diffuse ratios to actual physical diffuse amounts
    xDiffuse = xDiffuseRatio * curDIBValues.Width
    yDiffuse = yDiffuseRatio * curDIBValues.Height
    
    'These values will help us access locations in the array more quickly.
    Dim xStrideDiffuseX As Long, xStrideDiffuseY As Long
        
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()

    'pdRandomize handles random number duties
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_AutomaticAndRandom
    
    'hDX and hDY are the half-values (or radius) of the diffuse area.  Pre-calculating them is faster than recalculating
    ' them every time we need to access a radius value.
    Dim hDX As Double, hDY As Double
    hDX = xDiffuse / 2
    hDY = yDiffuse / 2
    
    'Finally, these two variables will be used to store the position of diffused pixels
    Dim diffuseX As Long, diffuseY As Long
    
    'Loop through each pixel in the image, diffusing as we go
    For y = initY To finalY
    For x = initX To finalX
        
        diffuseX = cRandom.GetRandomFloat_WH() * xDiffuse - hDX
        diffuseY = cRandom.GetRandomFloat_WH() * yDiffuse - hDY
        
        xStrideDiffuseX = diffuseX + x
        xStrideDiffuseY = diffuseY + y
            
        'Make sure the diffused pixel is within image boundaries, and if not adjust it according to the user's
        ' "wrapPixels" setting.
        If wrapPixels Then
            If (xStrideDiffuseX < initX) Then xStrideDiffuseX = xStrideDiffuseX + finalX
            If (xStrideDiffuseY < initY) Then xStrideDiffuseY = xStrideDiffuseY + finalY
            
            If (xStrideDiffuseX > finalX) Then xStrideDiffuseX = xStrideDiffuseX - finalX
            If (xStrideDiffuseY > finalY) Then xStrideDiffuseY = xStrideDiffuseY - finalY
        Else
            If (xStrideDiffuseX < initX) Then xStrideDiffuseX = initX
            If (xStrideDiffuseY < initY) Then xStrideDiffuseY = initY
            
            If (xStrideDiffuseX > finalX) Then xStrideDiffuseX = finalX
            If (xStrideDiffuseY > finalY) Then xStrideDiffuseY = finalY
        End If
        
        dstImageData(x, y) = srcImageData(xStrideDiffuseX, xStrideDiffuseY)
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate all image arrays
    workingDIB.UnwrapLongArrayFromDIB dstImageData
    srcDIB.UnwrapLongArrayFromDIB srcImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
     
End Sub

Private Sub sltX_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then DiffuseCustom GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltY_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "xsize", sltX.Value
        .AddParam "ysize", sltY.Value
        .AddParam "wrap", chkWrap.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
