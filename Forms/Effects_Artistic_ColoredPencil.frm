VERSION 5.00
Begin VB.Form FormPencil 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Colored pencil"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   360
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
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdDropDown cboStyle 
      Height          =   735
      Left            =   6000
      TabIndex        =   4
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "style"
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
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "tip radius"
      Min             =   1
      Max             =   100
      Value           =   3
      DefaultValue    =   3
   End
   Begin PhotoDemon.pdSlider sltIntensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3600
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "pressure"
      Max             =   300
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
End
Attribute VB_Name = "FormPencil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Pencil Sketch Image Effect
'Copyright 2001-2026 by Tanner Helland
'Created: sometime 2001
'Last updated: 26/July/17
'Last update: performance improvements, migrate to XML params
'
'PhotoDemon has provided a pencil sketch tool for a long time, but despite going through many incarnations, it always
' used low-quality, "quick and dirty" approximations.
'
'In July '14, this changed, and the entire tool was rethought from the ground up.  A dialog is now provided, with options
' for pencil style, tip thickness, and stroke pressure.  This yields much more flexible results, and the use of PD's
' central compositor for overlaying various image copies keeps things nice and fast.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To improve performance, we cache a local temporary DIB when generating previews
Private m_blurDIB As pdDIB

'Apply a "colored pencil" effect to an image
'Inputs:
' 1) radius of the pencil tip (min 1, no real max - but processing speed obviously drops as the radius increases)
' 2) color intensity, which controls the vibrance applied to the resulting color
' 3) pencil style, a nebulous setting that controls blend mode and post-processing, among other items.  Current values include:
'    0 - normal
'    1 - luminous
'    2 - pastel
'    3 - graphite
Public Sub fxColoredPencil(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Sketching image with pencils..."
    
    'Parse parameters out of the incoming param string
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim penRadius As Long, colorIntensity As Double, pencilStyle As Long
    
    With cParams
        penRadius = .GetLong("radius", sltRadius.Value)
        colorIntensity = .GetDouble("intensity", sltIntensity.Value)
        pencilStyle = .GetLong("style", cboStyle.ListIndex)
    End With
    
    'Reverse the intensity input; this way, positive values make the image more vibrant.  Negative values make it less vibrant.
    ' Note that the adjustment also varies by pencil style; typically it's used as a vibrance adjustment, but in some modes,
    ' we switch it out for gamma or contrast control.
    Select Case pencilStyle
    
        Case 0, 1
            colorIntensity = -0.01 * colorIntensity
            
        Case 2, 3
            colorIntensity = (302 - (colorIntensity + 101)) / 300
            
    End Select
    
    'More color variables - in this case, sums for each color component
    Dim r As Long, g As Long, b As Long, maxVal As Long
    Dim amtVal As Double, avgVal As Double
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, doNotUnPremultiplyAlpha:=True
    
    'Create a copy of the image.  "Colored pencil" requires a blurred image copy as part of the effect, and we maintain
    ' that copy separate from the original (as the two must be blended as the final step of the filter).
    If (m_blurDIB Is Nothing) Then Set m_blurDIB = New pdDIB
    m_blurDIB.CreateFromExistingDIB workingDIB
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        penRadius = penRadius * curDIBValues.previewModifier
        
    'If this is not a preview, initialize the main program progress bar
    Else
        SetProgBarMax finalY * 3 + finalX * 5
        Dim progBarCheck As Long
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Enforce a minimum radius amount
    If (penRadius < 2) Then penRadius = 2
    
    'Start by creating the blurred DIB
    If CreateApproximateGaussianBlurDIB(penRadius, workingDIB, m_blurDIB, 3, toPreview, finalY * 3 + finalX * 5) Then
        
        Dim progBarOffset As Long
        progBarOffset = finalY * 3 + finalX * 3
        
        'Now that we have a gaussian DIB created in blurDIB, we can point arrays toward it and the source DIB
        Dim srcImageData() As Byte, srcSA As SafeArray2D
        m_blurDIB.SetAlphaPremultiplication False
        m_blurDIB.WrapArrayAroundDIB srcImageData, srcSA
        
        Dim xStride As Long
                
        'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
        Dim grayLookUp() As Byte
        ReDim grayLookUp(0 To 765) As Byte
        
        For x = 0 To 765
            grayLookUp(x) = x \ 3
        Next x
                
        'Invert the source DIB, and optionally, apply grayscale as well
        For x = initX To finalX
            xStride = x * 4
        For y = initY To finalY
            
            b = srcImageData(xStride, y)
            g = srcImageData(xStride + 1, y)
            r = srcImageData(xStride + 2, y)
            
            'Normally, we invert the raw pixel data only...
            If (pencilStyle <> 1) Then
                srcImageData(xStride, y) = 255 - b
                srcImageData(xStride + 1, y) = 255 - g
                srcImageData(xStride + 2, y) = 255 - r
                
            '...but for the "luminous" color mode, we also convert the image to grayscale
            Else
                g = 255 - grayLookUp(r + g + b)
                srcImageData(xStride, y) = g
                srcImageData(xStride + 1, y) = g
                srcImageData(xStride + 2, y) = g
            End If
            
        Next y
            If (Not toPreview) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal progBarOffset + x
                End If
            End If
        Next x
            
        'Release our array copy
        m_blurDIB.UnwrapArrayFromDIB srcImageData
        
        'Apply premultiplication to both layers prior to compositing.  (Note that the working DIB
        ' is *already* in premultiplied space.)
        m_blurDIB.SetAlphaPremultiplication True
        
        'A pdCompositor class will help us blend the invert+blur image back onto the original image
        Dim cComposite As pdCompositor
        Set cComposite = New pdCompositor
        
        'Composite our invert+blur image against the base layer (workingDIB) using the COLOR DODGE blend mode;
        ' this will emphasize areas where the layers differ, while ignoring areas where they're the same.
        Dim topBlendMode As PD_BlendMode
        If (pencilStyle <> 2) Then topBlendMode = BM_ColorDodge Else topBlendMode = BM_LinearDodge
        cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, m_blurDIB, topBlendMode
        
        'Remove premultiplied alpha from the result
        workingDIB.SetAlphaPremultiplication False
        
        'Release any temporary DIBs as they are no longer required
        If (Not toPreview) Then Set m_blurDIB = Nothing
        
        'Some modes requires post-production gamma correction.  Build a lookup table now.
        Dim gammaTable() As Byte
        ReDim gammaTable(0 To 255) As Byte
        
        If (pencilStyle = 2) Or (pencilStyle = 3) Then
        
            Dim tmpVal As Double
            
            For x = 0 To 255
                tmpVal = x / 255
                tmpVal = tmpVal ^ (1# / colorIntensity)
                tmpVal = tmpVal * 255
                
                If (tmpVal > 255) Then
                    tmpVal = 255
                ElseIf (tmpVal < 0) Then
                    tmpVal = 0
                End If
                
                gammaTable(x) = tmpVal
            Next x
        
        End If
        
        'Point our byte array at workingDIB, so we can apply a final vibrance pass using the specified color intensity
        workingDIB.WrapArrayAroundDIB srcImageData, srcSA
        
        progBarOffset = finalY * 3 + finalX * 4
        
        'Adjust vibrance
        For x = initX To finalX
            xStride = x * 4
        For y = initY To finalY
            
            b = srcImageData(xStride, y)
            g = srcImageData(xStride + 1, y)
            r = srcImageData(xStride + 2, y)
                
            'Calculate the gray value using different methods for each pencil style
            If (pencilStyle = 0) Or (pencilStyle = 1) Then
            
                avgVal = grayLookUp(r + g + b)
                maxVal = Max3Int(r, g, b)
                
                'Calculate a vibrance-adjusted average, using the gray as our base
                amtVal = ((Abs(maxVal - avgVal) / 127) * colorIntensity)
                
                If (r <> maxVal) Then
                    r = r + (maxVal - r) * amtVal
                    If (r < 0) Then r = 0
                    If (r > 255) Then r = 255
                End If
                
                If (g <> maxVal) Then
                    g = g + (maxVal - g) * amtVal
                    If (g < 0) Then g = 0
                    If (g > 255) Then g = 255
                End If
                
                If (b <> maxVal) Then
                    b = b + (maxVal - b) * amtVal
                    If (b < 0) Then b = 0
                    If (b > 255) Then b = 255
                End If
                    
            ElseIf (pencilStyle = 2) Then
            
                r = gammaTable(r)
                g = gammaTable(g)
                b = gammaTable(b)
            
            'At present, the only other possibility is pencilStyle = 3
            Else
                r = gammaTable(grayLookUp(r + g + b))
                g = r
                b = r
            End If
            
            srcImageData(xStride, y) = b
            srcImageData(xStride + 1, y) = g
            srcImageData(xStride + 2, y) = r
            
        Next y
            If (Not toPreview) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal progBarOffset + x
                End If
            End If
        Next x
        
        'Release our array once more
        workingDIB.UnwrapArrayFromDIB srcImageData
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub cboStyle_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Colored pencil", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully loaded
    cmdBar.SetPreviewStatus False
    
    'Populate the style drop-down
    cboStyle.Clear
    cboStyle.AddItem "Normal", 0
    cboStyle.AddItem "Luminous", 1
    cboStyle.AddItem "Pastel", 2
    cboStyle.AddItem "Graphite", 3
    
    cboStyle.ListIndex = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltIntensity_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

'Render a new effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then fxColoredPencil GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "radius", sltRadius.Value
        .AddParam "intensity", (sltIntensity.Value - 100)
        .AddParam "style", cboStyle.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
