VERSION 5.00
Begin VB.Form FormPanAndZoom 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Pan and zoom"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltHorizontal 
      Height          =   705
      Left            =   6000
      TabIndex        =   1
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "horizontal pan"
      Min             =   -64
      Max             =   64
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltVertical 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "vertical pan"
      Min             =   -64
      Max             =   64
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltZoom 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "zoom"
      Min             =   -10
      SigDigits       =   2
   End
   Begin PhotoDemon.pdSlider sltQuality 
      Height          =   705
      Left            =   6000
      TabIndex        =   6
      Top             =   3480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "quality"
      Min             =   1
      Max             =   5
      Value           =   2
      NotchPosition   =   2
      NotchValueCustom=   2
   End
   Begin PhotoDemon.pdDropDown cboEdges 
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   4680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "if pixels lie outside the image..."
   End
End
Attribute VB_Name = "FormPanAndZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Pan and Zoom Effect Interface
'Copyright 2013-2017 by Tanner Helland
'Created: 28/May/13
'Last updated: 27/July/17
'Last update: performance improvements, migrate to XML params
'
'Dialog for handling a Ken Burns transform (http://en.wikipedia.org/wiki/Ken_burns_effect).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub cboEdges_Click()
    UpdatePreview
End Sub

'Apply a Ken Burns effect (basically, variable pan and zoom parameters with optional wrapping)
Public Sub PanAndZoomFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Applying pan and zoom (Ken Burns) effect..."
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    Dim hPan As Double, vPan As Double, newZoom As Double
    Dim edgeHandling As Long, superSamplingAmount As Long
    
    With cParams
        hPan = .GetDouble("horizontal", sltHorizontal.Value)
        vPan = .GetDouble("vertical", sltVertical.Value)
        newZoom = .GetDouble("zoom", sltZoom.Value)
        edgeHandling = .GetLong("edges", cboEdges.ListIndex)
        superSamplingAmount = .GetLong("quality", sltQuality.Value)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    PrepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'If this is a preview, adjust incoming parameters so they are representative of the final image
    If toPreview Then
        hPan = hPan * curDIBValues.previewModifier
        vPan = vPan * curDIBValues.previewModifier
    End If
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
                
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.SetDistortParameters qvDepth, edgeHandling, (superSamplingAmount <> 1), curDIBValues.maxX, curDIBValues.maxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = FindBestProgBarValue()
    
    '***************************************
    ' /* BEGIN SUPERSAMPLING PREPARATION */
    
    'Due to the way this filter works, supersampling yields much better results.  Because supersampling is extremely
    ' energy-intensive, this tool uses a sliding value for quality, as opposed to a binary TRUE/FALSE for antialiasing.
    ' (For all but the lowest quality setting, antialiasing will be used, and higher quality values will simply increase
    '  the amount of supersamples taken.)
    Dim newR As Long, newG As Long, newB As Long, newA As Long
    Dim r As Long, g As Long, b As Long, a As Long
    Dim tmpSum As Long, tmpSumFirst As Long
    Dim avgSamples As Double
    
    'Use the passed super-sampling constant (displayed to the user as "quality") to come up with a number of actual
    ' pixels to sample.  (The total amount of sampled pixels will range from 1 to 13).  Note that supersampling
    ' coordinates are precalculated and cached using a modified rotated grid function, which is consistent throughout PD.
    Dim numSamples As Long
    Dim ssX() As Single, ssY() As Single
    Filters_Area.GetSupersamplingTable superSamplingAmount, numSamples, ssX, ssY
    
    'Because supersampling will be used in the inner loop as (samplecount - 1), permanently decrease the sample
    ' count in advance.
    numSamples = numSamples - 1
    
    'Additional variables are needed for supersampling handling
    Dim j As Double, k As Double
    Dim sampleIndex As Long, numSamplesUsed As Long
    Dim superSampleVerify As Long, ssVerificationLimit As Long
    
    'Adaptive supersampling allows us to bypass supersampling if a pixel doesn't appear to benefit from it.  The superSampleVerify
    ' variable controls how many pixels are sampled before we perform an adaptation check.  At present, the rule is:
    ' Quality 3: check a minimum of 2 samples, Quality 4: check minimum 3 samples, Quality 5: check minimum 4 samples
    superSampleVerify = superSamplingAmount - 2
    
    'Alongside a variable number of test samples, adaptive supersampling requires some threshold that indicates samples
    ' are close enough that further supersampling is unlikely to improve output.  We calculate this as a minimum variance
    ' as 1.5 per channel (for a total of 6 variance per pixel), multiplied by the total number of samples taken.
    ssVerificationLimit = superSampleVerify * 6
    
    'To improve performance for quality 1 and 2 (which perform no supersampling), we can forcibly disable supersample checks
    ' by setting the verification checker to some impossible value.
    If (superSampleVerify <= 0) Then superSampleVerify = LONG_MAX
    
    ' /* END SUPERSAMPLING PREPARATION */
    '*************************************
    
    'Pan/zoom requires a number of specialized variables
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
    
    'Invert the vertical pan parameters, as that seems to be more intuitive to non-programmers
    vPan = -1# * vPan
    
    'Zoom is passed in as a value from -10 to 10.  0 implies no change.  We need to convert this
    ' value to something that can actually be used to modify zoom.
    If (newZoom >= 0#) Then
        newZoom = 1# / (newZoom + 1#)
    Else
        newZoom = -1# * (newZoom - 1#)
    End If
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curDIBValues.Width
    tHeight = curDIBValues.Height
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
                         
        'Reset all supersampling values
        newR = 0
        newG = 0
        newB = 0
        newA = 0
        numSamplesUsed = 0
        
        'Remap the coordinates around a center point of (0, 0)
        nX = x - midX
        nY = y - midY
        
        'Sample a number of source pixels corresponding to the user's supplied quality value; more quality means
        ' more samples, and much better representation in the final output.
        For sampleIndex = 0 To numSamples
            
            'Offset the pixel amount by the supersampling lookup table
            j = nX + ssX(sampleIndex)
            k = nY + ssY(sampleIndex)
            
            srcX = midX - hPan + j * newZoom
            srcY = midY - vPan + k * newZoom
            
            'Use the filter support class to interpolate and edge-wrap pixels as necessary
            fSupport.GetColorsFromSource r, g, b, a, srcX, srcY, srcImageData, x, y
            
            'If adaptive supersampling is active, apply the "adaptive" aspect.  Basically, calculate a variance for the currently
            ' collected samples.  If variance is low, assume this pixel does not require further supersampling.
            ' (Note that this is an ugly shorthand way to calculate variance, but it's fast, and the chance of false outliers is
            '  small enough to make it preferable over a true variance calculation.)
            If (sampleIndex = superSampleVerify) Then
                
                'Calculate variance for the first two pixels (Q3), three pixels (Q4), or four pixels (Q5)
                tmpSum = (r + g + b + a) * superSampleVerify
                tmpSumFirst = newR + newG + newB + newA
                
                'If variance is below 1.5 per channel per pixel, abort further supersampling
                If (Abs(tmpSum - tmpSumFirst) < ssVerificationLimit) Then Exit For
            
            End If
            
            'Increase the sample count
            numSamplesUsed = numSamplesUsed + 1
            
            'Add the retrieved values to our running averages
            newR = newR + r
            newG = newG + g
            newB = newB + b
            newA = newA + a
        
        Next sampleIndex
        
        'Find the average values of all samples, apply to the pixel, and move on!
        avgSamples = 1# / numSamplesUsed
        newR = newR * avgSamples
        newG = newG * avgSamples
        newB = newB * avgSamples
        newA = newA * avgSamples
        
        dstImageData(quickVal, y) = newB
        dstImageData(quickVal + 1, y) = newG
        dstImageData(quickVal + 2, y) = newR
        dstImageData(quickVal + 3, y) = newA
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'Safely deallocate all image arrays
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Pan and zoom", , GetLocalParamString(), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cboEdges.ListIndex = EDGE_WRAP
    sltQuality.Value = 2
End Sub

Private Sub Form_Load()

    'Suspend previews while we initialize all the controls
    cmdBar.MarkPreviewStatus False

    'Note the current image's width and height, which will be needed to adjust the preview effect
    Dim iWidth As Long, iHeight As Long
    If pdImages(g_CurrentImage).IsSelectionActive Then
        Dim selBounds As RECTF
        selBounds = pdImages(g_CurrentImage).MainSelection.GetBoundaryRect()
        iWidth = selBounds.Width
        iHeight = selBounds.Height
    Else
        iWidth = pdImages(g_CurrentImage).Width
        iHeight = pdImages(g_CurrentImage).Height
    End If
        
    sltHorizontal.Min = -1 * (iWidth \ 2)
    sltHorizontal.Max = iWidth \ 2
    sltVertical.Min = -1 * (iHeight \ 2)
    sltVertical.Max = iHeight \ 2
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cboEdges, EDGE_WRAP
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then PanAndZoomFilter GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltQuality_Change()
    UpdatePreview
End Sub

Private Sub sltVertical_Change()
    UpdatePreview
End Sub

Private Sub sltHorizontal_Change()
    UpdatePreview
End Sub

Private Sub sltZoom_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        .AddParam "horizontal", sltHorizontal.Value
        .AddParam "vertical", sltVertical.Value
        .AddParam "zoom", sltZoom.Value
        .AddParam "edges", cboEdges.ListIndex
        .AddParam "quality", sltQuality.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
