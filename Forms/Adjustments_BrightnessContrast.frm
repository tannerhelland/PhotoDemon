VERSION 5.00
Begin VB.Form FormBrightnessContrast 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Brightness and contrast"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12075
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
   ScaleWidth      =   805
   Begin PhotoDemon.pdButtonStrip btsModel 
      Height          =   975
      Left            =   6000
      TabIndex        =   5
      Top             =   2880
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "model"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdCheckBox chkSample 
      Height          =   330
      Left            =   6120
      TabIndex        =   3
      Top             =   3960
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   582
      Caption         =   "sample image for true contrast (slower but more accurate)"
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sltBright 
      Height          =   705
      Left            =   6000
      TabIndex        =   1
      Top             =   1200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "brightness"
      Min             =   -255
      Max             =   255
   End
   Begin PhotoDemon.pdSlider sltContrast 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "contrast"
      Min             =   -100
      Max             =   100
   End
End
Attribute VB_Name = "FormBrightnessContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Brightness and Contrast Handler
'Copyright 2001-2026 by Tanner Helland
'Created: 2/6/01
'Last updated: 21/October/22
'Last update: fix potential overflow due to unintended use of \ instead of /
'             (see https://github.com/tannerhelland/PhotoDemon/issues/452)
'
'Basic brightness/contrast handler.  A legacy LUT-based method is provided, but the modern L*a*b* implementation
' (via LittleCMS) is preferred.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'While previewing, we don't need to repeatedly sample contrast.  Just do it once and store the value.
Private m_previewHasSampled As Boolean
Private m_previewSampledContrast As Long

Private Sub btsModel_Click(ByVal buttonIndex As Long)
    SetLegacyVisibility
    UpdatePreview
End Sub

Private Sub SetLegacyVisibility()
    chkSample.Visible = (btsModel.ListIndex <> 0)
End Sub

'Update the preview when the "sample contrast" checkbox value is changed
Private Sub chkSample_Click()
    UpdatePreview
End Sub

'Single routine for modifying both brightness and contrast.  Brightness is in the range (-255,255) while
' contrast is (-100,100).  Optionally, the image can be sampled to obtain a true midpoint for the contrast function.
Public Sub BrightnessContrast(ByVal functionParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Start by extract individual function parameters from the XML string we're passed
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString functionParams
    
    Dim newBrightness As Long, newContrast As Double
    newBrightness = cParams.GetLong("brightness", 0)
    newContrast = cParams.GetLong("contrast", 0#)
    
    Dim useLegacyModel As Boolean, sampleContrast As Boolean
    useLegacyModel = cParams.GetBool("uselegacy", False)
    sampleContrast = cParams.GetBool("samplecontrast", False)
    
    If (Not toPreview) Then Message "Adjusting brightness and contrast..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim srcImageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStart As Long, xStop As Long
    xStart = initX * 4
    xStop = finalX * 4
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then
        ProgressBars.SetProgBarMax finalY
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'How we apply brightness varies; by default, a modern L*a*b*-based transform is used.  This produces a
    ' much higher-quality result, with significantly less clipping at either end of the histogram.
    
    '(LittleCMS is used for the transform, and if it's missing or disabled, we obviously can't proceed; if that happens,
    ' we fall back to the default brightness/contrast transform.
    Dim modernAlgorithmFailed As Boolean: modernAlgorithmFailed = True
    If (Not useLegacyModel) And PluginManager.IsPluginCurrentlyEnabled(CCP_LittleCMS) Then
        
        'Convert the incoming brightness and contrast values to ranges appropriate for L*a*b*
        Dim tmpBright As Double, tmpContrast As Double
        tmpBright = CDbl(newBrightness / 16#)
        tmpContrast = CDbl(newContrast / 800#) + 1#
        
        'We cheat and also use contrast as saturation, to flatten colors a bit when contrast is reduced
        Dim tmpSaturation As Double
        tmpSaturation = (newContrast / 12#)
        If (newContrast > 0) Then tmpSaturation = tmpSaturation + 1
        
        'Create an abstract LCMS transform that defines this adjustment
        Dim labTransform As pdLCMSTransform
        Set labTransform = New pdLCMSTransform
        If labTransform.CreateRGBModificationTransform(, tmpBright, tmpContrast, , tmpSaturation, , , INTENT_PRESERVE_K_PLANE_PERCEPTUAL) Then
            
            Dim scanWidthBytes As Long, scanWidthPixels As Long
            scanWidthBytes = workingDIB.GetDIBStride
            scanWidthPixels = workingDIB.GetDIBWidth
            
            'Apply the transform one line at a time
            For y = initY To finalY
                labTransform.ApplyTransformToArbitraryMemory workingDIB.GetDIBScanline(y), workingDIB.GetDIBScanline(y), scanWidthBytes, scanWidthBytes, 1, scanWidthPixels, False
                If (Not toPreview) Then
                    If (y And progBarCheck) = 0 Then
                        If Interface.UserPressedESC() Then Exit For
                        ProgressBars.SetProgBarVal y
                    End If
                End If
            Next y
            
            modernAlgorithmFailed = False
            
        End If
        
    End If
    
    'If the legacy mode is required (either by user choice or failure of LittleCMS), apply it now
    If (useLegacyModel Or modernAlgorithmFailed) And (Not g_cancelCurrentAction) Then
        
        Dim tmpSA1D As SafeArray1D
        
        Dim newBCTable(0 To 255) As Byte
        Dim btCalc As Long
        
        'Calculate brightness first; if no brightness change is being applied, no problem; the LUT will just
        ' be an identity LUT
        For x = 0 To 255
            btCalc = x + newBrightness
            If (btCalc > 255) Then btCalc = 255
            If (btCalc < 0) Then btCalc = 0
            newBCTable(x) = btCalc
        Next x
        
        If (newContrast <> 0) Then
        
            'Calculate contrast second.  Contrast is unique because it may require us to sample the source image
            ' to find the image's "true" luminance mean.
            Dim imgMean As Long
            
            'Sampled contrast is my invention; traditionally contrast pushes colors toward or away from gray.
            ' I like the option to push the colors toward or away from the image's actual midpoint, which
            ' may not be gray.  For most white-balanced photos the difference is minimal, but for images with
            ' non-traditional white balance, sampled contrast offers better results.
            If sampleContrast Then
                
                'During preview mode, we cache sampled contrast so we don't have to recalculate it on each redraw
                If (toPreview And m_previewHasSampled) Then
                    imgMean = m_previewSampledContrast
                Else
                
                    Dim rTotal As Single, gTotal As Single, bTotal As Single
                    rTotal = 0!
                    gTotal = 0!
                    bTotal = 0!
                    
                    Dim numOfPixels As Long
                    numOfPixels = 0
                    
                    For y = initY To finalY
                        workingDIB.WrapArrayAroundScanline srcImageData, tmpSA1D, y
                    For x = xStart To xStop Step 4
                        bTotal = bTotal + srcImageData(x)
                        gTotal = gTotal + srcImageData(x + 1)
                        rTotal = rTotal + srcImageData(x + 2)
                        numOfPixels = numOfPixels + 1
                    Next x
                    Next y
                    
                    rTotal = rTotal / numOfPixels
                    gTotal = gTotal / numOfPixels
                    bTotal = bTotal / numOfPixels
                    
                    imgMean = Int((rTotal + gTotal + bTotal) / 3! + 0.5!)
                    
                    'As mentioned earlier, cache the sample contrast during preview mode
                    If toPreview Then
                        m_previewSampledContrast = imgMean
                        m_previewHasSampled = True
                    End If
                
                End If
                    
            'If we're not using true contrast, set the mean to the traditional 127
            Else
                imgMean = 127
            End If
            
            'Use the calculated mean to complete the look-up table
            Dim ctCalc As Long, srcBrightness As Long
            For x = 0 To 255
                srcBrightness = newBCTable(x)
                ctCalc = srcBrightness + (((srcBrightness - imgMean) * newContrast) \ 100)
                If (ctCalc > 255) Then ctCalc = 255
                If (ctCalc < 0) Then ctCalc = 0
                newBCTable(x) = ctCalc
            Next x
        
        End If
        
        'Apply the LUT to the image!
        For y = initY To finalY
            workingDIB.WrapArrayAroundScanline srcImageData, tmpSA1D, y
        For x = xStart To xStop Step 4
            srcImageData(x) = newBCTable(srcImageData(x))
            srcImageData(x + 1) = newBCTable(srcImageData(x + 1))
            srcImageData(x + 2) = newBCTable(srcImageData(x + 2))
        Next x
            If (Not toPreview) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y
                End If
            End If
        Next y
        
        'With our work complete, point the local array away from the DIB
        workingDIB.UnwrapArrayFromDIB srcImageData
    
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic

End Sub

'OK button.  Note that the command bar class handles validation, form hiding, and form unload for us.
Private Sub cmdBar_OKClick()
    Process "Brightness and contrast", , GetFunctionParamString(), UNDO_Layer
End Sub

Private Function GetFunctionParamString() As String
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    With cParams
        .AddParam "brightness", sltBright.Value
        .AddParam "contrast", sltContrast.Value
        .AddParam "uselegacy", (btsModel.ListIndex = 1)
        .AddParam "samplecontrast", chkSample.Value
    End With
    GetFunctionParamString = cParams.GetParamString
End Function

'Sometimes the command bar will perform actions (like loading a preset) that require an updated preview.  This function
' is fired by the control when it's ready for such an update.
Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    m_previewHasSampled = 0
    m_previewSampledContrast = 0
    
    btsModel.AddItem "modern", 0
    btsModel.AddItem "legacy", 1
    btsModel.ListIndex = 0
    SetLegacyVisibility
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltBright_Change()
    UpdatePreview
End Sub

Private Sub sltContrast_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then BrightnessContrast GetFunctionParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub
