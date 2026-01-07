VERSION 5.00
Begin VB.Form FormImageCompare 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Compare images"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6510
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
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCheckBox chkSettings 
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "resize images as necessary"
   End
   Begin PhotoDemon.pdLabel lblSettings 
      Height          =   375
      Left            =   120
      Top             =   3240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      Caption         =   "options"
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   5055
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   1508
   End
   Begin PhotoDemon.pdDropDown ddSource 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1296
      Caption         =   "base image and layer"
   End
   Begin PhotoDemon.pdDropDown ddSource 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      FontSizeCaption =   11
   End
   Begin PhotoDemon.pdDropDown ddSource 
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1296
      Caption         =   "comparison image and layer"
   End
   Begin PhotoDemon.pdDropDown ddSource 
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      FontSizeCaption =   11
   End
   Begin PhotoDemon.pdCheckBox chkSettings 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "compare using Lab color space"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkSettings 
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   4440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "generate difference map as a new layer"
      Value           =   0   'False
   End
End
Attribute VB_Name = "FormImageCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Compare Images Dialog
'Copyright 2019-2026 by Tanner Helland
'Created: 19/November/19
'Last updated: 21/November/19
'Last update: wrap up initial build
'
'Sometimes it's helpful to compare two images mathematically (e.g. not doing a quick-and-dirty comparison
' like layering + difference blend mode).  I use this to ensure that various lossy optimizations/estimatinos
' for adjustments and effects do not negatively impact underlying image quality vs "perfect" implementations.
'
'At present, this dialog only reports PSNR values, which are not ideal...
' (see https://en.wikipedia.org/wiki/Peak_signal-to-noise_ratio#Performance_comparison)
'... so someday I may add options for better metrics, like SSIM...
' (see https://en.wikipedia.org/wiki/Structural_similarity)
'... but there's no planned completion date for this at present.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_OpenImageIDs As pdStack

'Compare two arbitrary layers from two arbitrary images.  All settings must be encoded in a param string.
Public Sub CompareImages(ByRef listOfParameters As String)
    
    Message "Analyzing image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString listOfParameters
    
    Dim srcImageID As Long, srcLayerIdx As Long
    Dim cmpImageID As Long, cmpLayerIdx As Long
    Dim allowResize As Boolean, useLab As Boolean
    Dim generateDiffMap As Boolean
    
    With cParams
        srcImageID = .GetLong("source-image-id", 0, True)
        srcLayerIdx = .GetLong("source-layer-idx", 0, True)
        cmpImageID = .GetLong("compare-image-id", 0, True)
        cmpLayerIdx = .GetLong("compare-layer-idx", 0, True)
        allowResize = .GetBool("force-size-match", True, True)
        useLab = Strings.StringsEqual(.GetString("color-space", "rgb", True), "lab", True)
        generateDiffMap = .GetBool("generate-difference-map", False, True)
    End With
    
    'Ensure all image and layer references are valid
    Dim srcImage As pdImage, cmpImage As pdImage
    Set srcImage = PDImages.GetImageByID(srcImageID)
    Set cmpImage = PDImages.GetImageByID(cmpImageID)
    If (srcImage Is Nothing) Or (cmpImage Is Nothing) Then Exit Sub
    
    Dim srcLayer As pdLayer, cmpLayer As pdLayer
    Set srcLayer = srcImage.GetLayerByIndex(srcLayerIdx, False)
    Set cmpLayer = cmpImage.GetLayerByIndex(cmpLayerIdx, False)
    If (srcLayer Is Nothing) Or (cmpLayer Is Nothing) Then Exit Sub
    
    'Comparisons may not be lossless; if the two targets are *not* the same size,
    ' the comparison layer may be resized to match the size of the base layer.
    ' (Similarly, if either layer has active non-destructive transforms, we'll need
    ' to create a temporary copy with committed changes prior to comparison.)
    Dim targetWidth As Long, targetHeight As Long
    
    'Start with the base layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.CopyExistingLayer srcLayer
    
    If srcLayer.AffineTransformsActive(True) Then
        tmpLayer.ConvertToNullPaddedLayer srcImage.Width, srcImage.Height, True
        tmpLayer.CropNullPaddedLayer
    End If
    
    Set srcLayer = tmpLayer
    targetWidth = srcLayer.GetLayerDIB.GetDIBWidth()
    targetHeight = srcLayer.GetLayerDIB.GetDIBHeight()
    
    'Repeat above steps for comparison layer, with the added step of resizing to
    ' match base layer dimensions (if necessary; we may also be cropping/enlarging it).
    Set tmpLayer = New pdLayer
    tmpLayer.CopyExistingLayer cmpLayer
    If cmpLayer.AffineTransformsActive(True) Then
        tmpLayer.ConvertToNullPaddedLayer cmpImage.Width, cmpImage.Height, True
        tmpLayer.CropNullPaddedLayer
    End If
    
    Set cmpLayer = tmpLayer
    
    If (cmpLayer.GetLayerDIB.GetDIBWidth <> targetWidth) Or (cmpLayer.GetLayerDIB.GetDIBHeight <> targetHeight) Then
        
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        
        'Resize into a temporary DIB
        If allowResize Then
            tmpDIB.CreateFromExistingDIB cmpLayer.GetLayerDIB, srcLayer.GetLayerDIB.GetDIBWidth, srcLayer.GetLayerDIB.GetDIBHeight, GP_IM_Bilinear
            
        'Crop/enlarge
        Else
            tmpDIB.CreateBlank targetWidth, targetHeight, 32, 0, 0
            GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, cmpLayer.GetLayerDIB.GetDIBWidth, cmpLayer.GetLayerDIB.GetDIBHeight, cmpLayer.GetLayerDIB.GetDIBDC, 0, 0, vbSrcCopy
        End If
        
        'Copy the DIB into a temporary layer object
        Set cmpLayer = New pdLayer
        cmpLayer.CopyExistingLayer tmpLayer
        cmpLayer.SetLayerDIB tmpDIB
        
    End If
    
    'We now have a source and comparison layer that are guaranteed to have matching sizes
    ' (via either crop/padding or resampling).  This makes comparisons trivial.
    Dim baseDIB As pdDIB, cmpDIB As pdDIB
    Set baseDIB = srcLayer.GetLayerDIB
    Set cmpDIB = cmpLayer.GetLayerDIB
    
    'Un-premultiply both images
    baseDIB.SetAlphaPremultiplication False
    cmpDIB.SetAlphaPremultiplication False
    
    Dim pxBase() As Byte, pxCmp() As Byte
    Dim baseSA As SafeArray1D, cmpSA As SafeArray1D
    
    Dim x As Long, y As Long, xFinal As Long, yFinal As Long
    yFinal = targetHeight - 1
    If useLab Then xFinal = (targetWidth - 1) * 3 Else xFinal = (targetWidth - 1) * 4
    
    Dim checkInterval As Long
    ProgressBars.SetProgBarMax yFinal
    checkInterval = ProgressBars.FindBestProgBarValue()
    
    'If using Lab color space, prep a transform using LittleCMS
    Dim labTransform As pdLCMSTransform, labBase() As Single, labCmp() As Single
    If useLab Then
        Set labTransform = New pdLCMSTransform
        labTransform.CreateRGBAToLabTransform , False, INTENT_PERCEPTUAL, 0&
        ReDim labBase(0 To targetWidth * 3 - 1) As Single
        ReDim labCmp(0 To targetWidth * 3 - 1) As Single
    End If
    
    Dim rTotal As Double, gTotal As Double, bTotal As Double
    Dim rTmp As Long, gTmp As Long, bTmp As Long, rTmp2 As Long, gTmp2 As Long, bTmp2 As Long
    Dim rTmpF As Double, gTmpF As Double, bTmpF As Double, rTmp2F As Double, gTmp2F As Double, bTmp2F As Double
    
    'Lab color space has weird, arbitrary max/min values that are (obviously) dependent on the
    ' color space of the current image.  We currently default to sRGB values; these will need to be
    ' revisited in the future if images are compared across non-uniform working spaces!
    
    'For sRGB...
    ' L is on the range [0, 100]
    ' a is on the range [-79.276268, 93.544746398926]
    ' b is on the range [-112.0311279296875, 93.392997741699]
    Const labAMin As Double = -79.276268
    'Const labAMax As Double = 93.544746398926
    Const labAScale As Double = 93.544746398926 + 79.276268 'labAMax + Abs(labAMin)
    
    Const labBMin As Double = -112.031127929688
    'Const labBMax As Double = 93.392997741699
    Const labBScale As Double = 93.392997741699 + 112.031127929688 'labBMax + Abs(labBMin)
    
    Dim labARatio As Double, labBRatio As Double
    labARatio = 100# / labAScale
    labBRatio = 100# / labBScale
    
    For y = 0 To yFinal
        
        'Lab comparison
        If useLab Then
            
            labTransform.ApplyTransformToScanline baseDIB.GetDIBPointer() + baseDIB.GetDIBStride * y, VarPtr(labBase(0)), targetWidth
            labTransform.ApplyTransformToScanline cmpDIB.GetDIBPointer() + cmpDIB.GetDIBStride * y, VarPtr(labCmp(0)), targetWidth
            
            For x = 0 To xFinal Step 3
                rTmpF = labBase(x)
                gTmpF = (labBase(x + 1) + labAMin) * labARatio
                bTmpF = (labBase(x + 2) + labBMin) * labBRatio
                rTmp2F = labCmp(x)
                gTmp2F = (labCmp(x + 1) + labAMin) * labARatio
                bTmp2F = (labCmp(x + 2) + labBMin) * labBRatio
                rTmpF = rTmpF - rTmp2F
                gTmpF = gTmpF - gTmp2F
                bTmpF = bTmpF - bTmp2F
                rTotal = rTotal + (rTmpF * rTmpF)
                gTotal = gTotal + (gTmpF * gTmpF)
                bTotal = bTotal + (bTmpF * bTmpF)
            Next x
            
        'RGB comparison
        Else
            
            baseDIB.WrapArrayAroundScanline pxBase, baseSA, y
            cmpDIB.WrapArrayAroundScanline pxCmp, cmpSA, y
            
            For x = 0 To xFinal Step 4
                bTmp = pxBase(x)
                gTmp = pxBase(x + 1)
                rTmp = pxBase(x + 2)
                bTmp2 = pxCmp(x)
                gTmp2 = pxCmp(x + 1)
                rTmp2 = pxCmp(x + 2)
                bTmp = bTmp - bTmp2
                gTmp = gTmp - gTmp2
                rTmp = rTmp - rTmp2
                bTotal = bTotal + (bTmp * bTmp)
                gTotal = gTotal + (gTmp * gTmp)
                rTotal = rTotal + (rTmp * rTmp)
            Next x
    
        End If
    
        'Periodic UI updates
        If ((y And checkInterval) = 0) Then ProgressBars.SetProgBarVal y
        
    Next y
    
    'Safely unwrap array aliases
    baseDIB.UnwrapArrayFromDIB pxBase
    cmpDIB.UnwrapArrayFromDIB pxCmp
    
    'If the caller wants us to generate a difference map, do that before displaying comparison results
    If generateDiffMap Then
        
        'Merge the comparison DIB onto the base DIB
        baseDIB.SetAlphaPremultiplication True
        cmpDIB.SetAlphaPremultiplication True
        
        Dim cCompositor As pdCompositor
        Set cCompositor = New pdCompositor
        cCompositor.QuickMergeTwoDibsOfEqualSize baseDIB, cmpDIB, BM_Difference
        
        'Ask the parent pdImage to create a new layer object
        Dim newLayerID As Long
        newLayerID = srcImage.CreateBlankLayer(srcImage.GetActiveLayerIndex)
        
        'Ask the new layer to copy the contents of the layer we are duplicating
        srcImage.GetLayerByID(newLayerID).CopyExistingLayer srcLayer
        srcImage.GetLayerByID(newLayerID).SetLayerName g_Language.TranslateMessage("difference between %1 and %2", srcLayer.GetLayerName, cmpLayer.GetLayerName)
        
        'Replace the (temporary) DIB it created with the merged DIB we just created
        srcImage.GetLayerByID(newLayerID).SetLayerDIB baseDIB
        
        'Make the duplicate layer the active layer
        srcImage.SetActiveLayerByID newLayerID
        
        'Notify the parent image that the entire image now needs to be recomposited
        srcImage.NotifyImageChanged UNDO_Image_VectorSafe
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
        'Render the new image to screen
        CanvasManager.ActivatePDImage srcImage.imageID, "comparison finished", True, UNDO_Image
        
    End If
    
    'Calculate MSE on a per-channel basis
    Dim numPixels As Double
    numPixels = 1# / CDbl(targetWidth * targetHeight)
    Dim rMSE As Double, gMSE As Double, bMSE As Double
    bMSE = bTotal * numPixels
    gMSE = gTotal * numPixels
    rMSE = rTotal * numPixels
    
    'From MSE, calculate PSNR.  (65025 = 255 * 255)
    Dim rPSNR As Double, gPSNR As Double, bPSNR As Double
    Dim maxPossibleValue As Double
    If useLab Then maxPossibleValue = 100# Else maxPossibleValue = 255#
    maxPossibleValue = maxPossibleValue * maxPossibleValue
    
    If (bMSE > 0#) Then bPSNR = 10# * PDMath.Log10(maxPossibleValue / bMSE) Else bPSNR = -1#
    If (gMSE > 0#) Then gPSNR = 10# * PDMath.Log10(maxPossibleValue / gMSE) Else gPSNR = -1#
    If (rMSE > 0#) Then rPSNR = 10# * PDMath.Log10(maxPossibleValue / rMSE) Else rPSNR = -1#
    
    Dim rPNSRtext As String, gPNSRtext As String, bPNSRtext As String, undefText As String
    
    'Use infinity char if available
    If OS.IsWin7OrLater() Then
        undefText = ChrW$(&H221E)
    Else
        undefText = g_Language.TranslateMessage("infinite")
    End If
    
    If (bPSNR >= 0#) Then bPNSRtext = Format$(bPSNR, "0.0000") Else bPNSRtext = undefText
    If (gPSNR >= 0#) Then gPNSRtext = Format$(gPSNR, "0.0000") Else gPNSRtext = undefText
    If (rPSNR >= 0#) Then rPNSRtext = Format$(rPSNR, "0.0000") Else rPNSRtext = undefText
    
    'Depending on the requested output, produce a result
    Dim msgFinal As pdString
    Set msgFinal = New pdString
    
    Dim colorSpaceText As String
    If useLab Then colorSpaceText = g_Language.TranslateMessage("Lab") Else colorSpaceText = g_Language.TranslateMessage("RGB")
    
    msgFinal.AppendLine g_Language.TranslateMessage("Peak signal-to-noise ratio (PSNR) between these images:")
    msgFinal.AppendLineBreak
    If useLab Then msgFinal.Append g_Language.TranslateMessage("L*:") Else msgFinal.Append g_Language.TranslateMessage("red:")
    msgFinal.Append " "
    msgFinal.AppendLine rPNSRtext
    If useLab Then msgFinal.Append g_Language.TranslateMessage("a*:") Else msgFinal.Append g_Language.TranslateMessage("green:")
    msgFinal.Append " "
    msgFinal.AppendLine gPNSRtext
    If useLab Then msgFinal.Append g_Language.TranslateMessage("b*:") Else msgFinal.Append g_Language.TranslateMessage("blue:")
    msgFinal.Append " "
    msgFinal.AppendLine bPNSRtext
    
    msgFinal.AppendLineBreak
    If (bPSNR < 0#) And (gPSNR < 0#) And (rPSNR < 0#) Then
        msgFinal.Append g_Language.TranslateMessage("These images are identical.")
    Else
        
        Dim ratioSimilarity As Double
        
        'Lab is already perceptually uniform, so we don't need to mess with weights like we do with RGB
        If useLab Then
            ratioSimilarity = 1# - (Sqr(bTotal * numPixels) + Sqr(gTotal * numPixels) + Sqr(rTotal * numPixels)) / (Sqr(maxPossibleValue) * 3#)
        
        'Use luminance weighting to improve perceptual relevance.
        Else
            ratioSimilarity = 1# - (Sqr(bTotal * numPixels) * 0.0722 + Sqr(gTotal * numPixels) * 0.7152 + Sqr(rTotal * numPixels)) * 0.2126 / Sqr(maxPossibleValue)
        End If
        
        'Manually clip to avoid automatic rounding to 100.0% when images are very similar
        ratioSimilarity = ratioSimilarity * 100#
        If (ratioSimilarity < 100#) And (ratioSimilarity > 99.99) Then ratioSimilarity = 99.99
        msgFinal.Append g_Language.TranslateMessage("As a rough estimate, these images are %1% similar.", Format$(ratioSimilarity, "0.00"))
        
    End If
    
    Message "Finished."
    ProgressBars.ReleaseProgressBar
    
    PDMsgBox msgFinal.ToString(), vbOKOnly Or vbInformation Or vbApplicationModal, "Image comparison"
    
End Sub

Private Sub cmdBar_OKClick()
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "source-image-id", m_OpenImageIDs.GetInt(ddSource(0).ListIndex)
        .AddParam "source-layer-idx", (ddSource(1).ListCount - 1) - ddSource(1).ListIndex   'Layers are displayed in visual order
        .AddParam "compare-image-id", m_OpenImageIDs.GetInt(ddSource(2).ListIndex)
        .AddParam "compare-layer-idx", (ddSource(3).ListCount - 1) - ddSource(3).ListIndex   'Layers are displayed in visual order
        .AddParam "force-size-match", chkSettings(0).Value
        If chkSettings(1).Value Then .AddParam "color-space", "lab" Else .AddParam "color-space", "rgb"
        .AddParam "generate-difference-map", chkSettings(2).Value
    End With
    
    Me.Visible = False
    
    'If we are generating a new comparison layer, we need to create undo data
    If chkSettings(2).Value Then
        Process "Compare similarity", , cParams.GetParamString(), UNDO_Image
    Else
        Process "Compare similarity", , cParams.GetParamString(), UNDO_Nothing
    End If
    
End Sub

Private Sub ddSource_Click(Index As Integer)
    
    Select Case Index
        
        'Base image / Comparison image
        Case 0, 2
            PopulateLayerList Index
            
    End Select
    
End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()

    'Populate both drop-downs with a list of open images
    PDImages.GetListOfActiveImageIDs m_OpenImageIDs
    
    ddSource(0).SetAutomaticRedraws False
    ddSource(2).SetAutomaticRedraws False
    
    Dim srcLayerName As String, idxActiveImage As Long
    Dim i As Long
    
    For i = 0 To m_OpenImageIDs.GetNumOfInts - 1
        If (m_OpenImageIDs.GetInt(i) = PDImages.GetActiveImageID) Then idxActiveImage = i
        srcLayerName = Interface.GetWindowCaption(PDImages.GetImageByID(m_OpenImageIDs.GetInt(i)), False, True)
        ddSource(0).AddItem srcLayerName
        ddSource(2).AddItem srcLayerName
    Next i
    
    'Auto-select the currently active image as the source, and if another image is available,
    ' set it as the comparison object
    ddSource(0).ListIndex = idxActiveImage
    
    Dim cmpIndex As Long
    If (ddSource(0).ListCount > 1) Then
        cmpIndex = idxActiveImage + 1
        If (cmpIndex >= ddSource(0).ListCount) Then cmpIndex = idxActiveImage - 1
    Else
        cmpIndex = idxActiveImage
    End If
    
    ddSource(2).ListIndex = cmpIndex
    
    ddSource(0).SetAutomaticRedraws True
    ddSource(2).SetAutomaticRedraws True
    
    'Select an active layer from both drop-downs
    PopulateLayerList 0, True
    PopulateLayerList 2, True
    
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub PopulateLayerList(ByVal srcDropDown As Long, Optional ByVal isInitPopulator As Boolean = False)
    
    Dim ddTarget As Long
    
    'Source image vs comparison image
    If (srcDropDown = 0) Then ddTarget = 1 Else ddTarget = 3
    
    'Clear existing layer lists
    ddSource(ddTarget).SetAutomaticRedraws False
    ddSource(ddTarget).Clear
    
    Dim srcImage As pdImage
    Set srcImage = PDImages.GetImageByID(m_OpenImageIDs.GetInt(ddSource(srcDropDown).ListIndex))
    
    'Populate with layer names (in *descending* order)
    Dim i As Long
    For i = srcImage.GetNumOfLayers - 1 To 0 Step -1
        ddSource(ddTarget).AddItem srcImage.GetLayerByIndex(i).GetLayerName()
    Next i
    
    'Auto-select the currently active layer in either image, unless the two images are identical;
    ' in that case, attempt to select a neighboring layer (if one exists)
    If (srcDropDown = 0) Then
        ddSource(ddTarget).ListIndex = (srcImage.GetNumOfLayers - 1) - srcImage.GetActiveLayerIndex
    Else
        
        'Only one image exists; try to select different layers
        If (ddSource(0).ListIndex = ddSource(2).ListIndex) Then
            
            Dim idxLayer As Long
            idxLayer = ddSource(1).ListIndex
            If (srcImage.GetNumOfLayers > 0) Then
                idxLayer = idxLayer + 1
                If (idxLayer >= ddSource(ddTarget).ListCount) Then
                    idxLayer = ddSource(1).ListIndex - 1
                    If (idxLayer < 0) Then idxLayer = 0
                End If
            End If
            
            ddSource(ddTarget).ListIndex = idxLayer
            
        'Two different images exist; use active layer from both
        Else
            ddSource(ddTarget).ListIndex = (srcImage.GetNumOfLayers - 1) - srcImage.GetActiveLayerIndex
        End If
        
    End If
    
    ddSource(ddTarget).SetAutomaticRedraws True, True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
