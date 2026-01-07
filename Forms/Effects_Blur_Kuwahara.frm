VERSION 5.00
Begin VB.Form FormKuwahara 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Kuwahara"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11640
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
   ScaleWidth      =   776
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2280
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "radius"
      Min             =   1
      Value           =   1
      DefaultValue    =   1
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
End
Attribute VB_Name = "FormKuwahara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Kuwahara Blur Dialog
'Copyright 2019-2026 by Tanner Helland
'Created: 15/November/19
'Last updated: 19/November/19
'Last update: wrap up initial build
'
'Per Wikipedia (https://en.wikipedia.org/wiki/Kuwahara_filter):
' "The Kuwahara filter is a non-linear smoothing filter used in image processing for adaptive noise reduction."
'
'For performance and quality reasons, PhotoDemon's implementation calculates variance using luminance;
' the quadrant with the smallest luminance variance is then used for calculating average color.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub KuwaharaFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim radius As Long
    radius = cParams.GetLong("radius", sltRadius.Value)
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim dstImageData() As Byte, dstSA As SafeArray2D, dstSA1D As SafeArray1D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Reduce radius for previews.  (Note that this has to happen *after* the EffectPrep call, above,
    ' as that calculates the preview ratio for us.)
    If toPreview Then radius = radius * curDIBValues.previewModifier
    If (radius < 1) Then radius = 1
    
    If (Not toPreview) Then Message "Applying Kuwahara filter..."
    
    'Sums (and sums^2) for each of the 4 quadrants
    Dim sum0 As Long, sum1 As Long, sum2 As Long, sum3 As Long
    Dim sums0 As Long, sums1 As Long, sums2 As Long, sums3 As Long
    
    'Pixel count in each quadrant
    Dim numPixels As Long
    numPixels = (radius + 1) * (radius + 1)
    
    Dim invNumPixels As Double
    invNumPixels = 1# / CDbl(numPixels)
    
    'Pad the working DIB (so we don't have to deal with boundary checking)
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    Dim tmpDIB As pdDIB
    Filters_Layers.PadDIBClampedPixels radius, radius, workingDIB, tmpDIB
    tmpDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    'Instead of analyzing RGB data, we want to analyze luminance data alone.  The quadrant with
    ' the largest *luminance* variance is the one we'll use for RGB data as well.
    Dim grayData() As Byte
    DIBs.GetDIBGrayscaleMap tmpDIB, grayData, False
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = radius
    initY = radius
    finalX = radius + workingDIB.GetDIBWidth - 1
    finalY = radius + workingDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then
        ProgressBars.SetProgBarMax finalY
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim xOffset As Long
    Dim xMin As Long, xMax As Long, yMin As Long, yMax As Long
    Dim tstVariance As Double, vMin As Double
    
    Dim i As Long, j As Long
    Dim r As Long, g As Long, b As Long
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        
        'Destination pixels are applied on a clean line-by-line basis, so we can simply point
        ' at the current scanline.
        workingDIB.WrapArrayAroundScanline dstImageData, dstSA1D, y - radius
        
    For x = initX To finalX
    
        'Clear all trackers
        sum0 = 0
        sum1 = 0
        sum2 = 0
        sum3 = 0
        sums0 = 0
        sums1 = 0
        sums2 = 0
        sums3 = 0
        
        'Calculate variance for each quadrant
        xMax = x + radius
        yMax = y + radius
        
        For j = y To yMax
        For i = x To xMax
            
            'These could be calculated in a loop, but I manually unroll them for improved perf
            
            'Use non-branching variance calculation
            g = grayData(i - radius, j - radius)
            sum0 = sum0 + g
            sums0 = sums0 + g * g
            
            'Repeat for other three quadrants
            g = grayData(i, j - radius)
            sum1 = sum1 + g
            sums1 = sums1 + g * g
            
            g = grayData(i - radius, j)
            sum2 = sum2 + g
            sums2 = sums2 + g * g
            
            g = grayData(i, j)
            sum3 = sum3 + g
            sums3 = sums3 + g * g
            
        Next i
        Next j
        
        'Find the quadrant with lowest variance.  (Again, this could be done in a loop; I unroll for perf benefits.)
        tstVariance = sums0 - sum0 * sum0 * invNumPixels
        vMin = tstVariance
        xMin = x - radius
        yMin = y - radius
        
        tstVariance = sums1 - sum1 * sum1 * invNumPixels
        If (tstVariance < vMin) Then
            vMin = tstVariance
            xMin = x
        End If
        
        tstVariance = sums2 - sum2 * sum2 * invNumPixels
        If (tstVariance < vMin) Then
            vMin = tstVariance
            xMin = x - radius
            yMin = y
        End If
        
        tstVariance = sums3 - sum3 * sum3 * invNumPixels
        If (tstVariance < vMin) Then
            xMin = x
            yMin = y
        End If
        
        'Assign average values of the lowest quadrant
        xMin = xMin * 4
        xMax = xMin + radius * 4
        yMax = yMin + radius
        
        r = 0
        g = 0
        b = 0
        
        For j = yMin To yMax
        For i = xMin To xMax Step 4
            b = b + srcImageData(i, j)
            g = g + srcImageData(i + 1, j)
            r = r + srcImageData(i + 2, j)
        Next i
        Next j
        
        xOffset = (x - radius) * 4
        dstImageData(xOffset) = b \ numPixels
        dstImageData(xOffset + 1) = g \ numPixels
        dstImageData(xOffset + 2) = r \ numPixels
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate arrays
    tmpDIB.UnwrapArrayFromDIB srcImageData
    Set tmpDIB = Nothing
    workingDIB.UnwrapArrayFromDIB dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Kuwahara filter", , GetLocalParamString(), UNDO_Layer
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

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.KuwaharaFilter GetLocalParamString(), True, pdFxPreview
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
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
