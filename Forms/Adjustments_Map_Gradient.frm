VERSION 5.00
Begin VB.Form FormGradientMap 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Gradient map"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12120
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
   ScaleWidth      =   808
   Begin PhotoDemon.pdGradientSelector grdSource 
      Height          =   1215
      Left            =   6000
      TabIndex        =   4
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2143
      Caption         =   "gradient"
   End
   Begin PhotoDemon.pdDropDown cboBlendMode 
      Height          =   975
      Left            =   6000
      TabIndex        =   3
      Top             =   3480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "blend mode"
   End
   Begin PhotoDemon.pdSlider sltIntensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2520
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "intensity"
      Max             =   100
      Value           =   50
      DefaultValue    =   50
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
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   1323
   End
End
Attribute VB_Name = "FormGradientMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gradient Map Adjustment Dialog
'Copyright 2022-2022 by Tanner Helland
'Created: 03/March/2022
'Last updated: 03/March/2022
'Last update: initial build
'
'Gradient mapping is more useful as an adjustment layer, but until PD supports 'em
' this dialog will need to suffice.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To improve preview performance, a persistent preview DIB is cached locally
Private m_EffectDIB As pdDIB

'Apply a gradient map effect to the active image/layer
Public Sub ApplyGradientMap(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Applying gradient map..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim srcGradient As pd2DGradient
    Dim fxIntensity As Double, fxBlend As PD_BlendMode
    
    With cParams
        fxBlend = .GetLong("blendmode", BM_Normal)
        fxIntensity = .GetDouble("intensity", 50#)
        Set srcGradient = New pd2DGradient
        srcGradient.CreateGradientFromString .GetString("source-gradient", vbNullString)
    End With
    
    'Pull a LUT from the source gradient
    Dim palColors() As Long
    srcGradient.GetLookupTable palColors, 256
    
    'Create a local array and point it at the pixel data we want to operate on,
    ' and note that a 1D array works fine (scanlines don't matter)
    Dim srcPixels() As Byte, dstPixels() As Long
    Dim tmpSA As SafeArray2D, tmpSA1D As SafeArray1D, tmpSALong As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    'Create a copy of the working data
    If (m_EffectDIB Is Nothing) Then Set m_EffectDIB = New pdDIB
    m_EffectDIB.CreateFromExistingDIB workingDIB
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long, grayVal As Long
    Dim xOffset As Long
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        m_EffectDIB.WrapArrayAroundScanline srcPixels, tmpSA1D, y
        m_EffectDIB.WrapLongArrayAroundScanline dstPixels, tmpSALong, y
    For x = initX To finalX
        
        xOffset = x * 4
        
        'Get the source pixel color values and calculate a gray value
        b = srcPixels(xOffset)
        g = srcPixels(xOffset + 1)
        r = srcPixels(xOffset + 2)
        grayVal = (218 * r + 732 * g + 74 * b) \ 1024
        
        'Assign the lookup value to the Long-type array alias
        dstPixels(x) = palColors(grayVal)
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate all aliases
    m_EffectDIB.UnwrapArrayFromDIB srcPixels
    m_EffectDIB.UnwrapLongArrayFromDIB dstPixels
    
    'Merge the results
    m_EffectDIB.SetAlphaPremultiplication True
    workingDIB.SetAlphaPremultiplication True
    
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, m_EffectDIB, fxBlend, fxIntensity
    
    'On non-previews, free our intermediate copy
    If (Not toPreview) Then Set m_EffectDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub cboBlendMode_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Gradient map", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cboBlendMode.ListIndex = BM_Normal
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal
    
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub grdSource_GradientChanged()
    UpdatePreview
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
    If cmdBar.PreviewsAllowed Then Me.ApplyGradientMap GetLocalParamString(), True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "blendmode", cboBlendMode.ListIndex
        .AddParam "intensity", sltIntensity.Value
        .AddParam "source-gradient", grdSource.Gradient
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
