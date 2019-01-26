VERSION 5.00
Begin VB.Form FormMonochrome 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Monochrome Conversion"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12150
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   810
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdSlider sldDitheringAmount 
      Height          =   855
      Left            =   6000
      TabIndex        =   8
      Top             =   2400
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1508
      Caption         =   "dithering amount"
      Max             =   100
      Value           =   100
      DefaultValue    =   100
   End
   Begin PhotoDemon.pdDropDown cboDither 
      Height          =   855
      Left            =   6000
      TabIndex        =   7
      Top             =   1440
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1508
      Caption         =   "dithering"
   End
   Begin PhotoDemon.pdButtonStrip btsTransparency 
      Height          =   1065
      Left            =   6000
      TabIndex        =   6
      Top             =   4560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1879
      Caption         =   "transparency"
   End
   Begin PhotoDemon.pdSlider sltThreshold 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   1244
      Caption         =   "threshold"
      Min             =   1
      Max             =   254
      Value           =   127
      NotchPosition   =   2
      NotchValueCustom=   127
   End
   Begin PhotoDemon.pdCheckBox chkAutoThreshold 
      Height          =   330
      Left            =   6120
      TabIndex        =   4
      Top             =   930
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   582
      Caption         =   "automatically calculate threshold"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdColorSelector csMono 
      Height          =   615
      Index           =   0
      Left            =   6120
      TabIndex        =   1
      Top             =   3720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      curColor        =   0
   End
   Begin PhotoDemon.pdColorSelector csMono 
      Height          =   615
      Index           =   1
      Left            =   9120
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   6000
      Top             =   3360
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   503
      Caption         =   "final colors"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormMonochrome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Monochrome Conversion Form
'Copyright 2002-2019 by Tanner Helland
'Created: some time 2002
'Last updated: 02/April/18
'Last update: add "single neighbor" as a dithering option
'
'The meat of this form is in the module with the same name...look there for
' real algorithm info.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsTransparency_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cboDither_Click()
    If chkAutoThreshold.Value Then sltThreshold.Value = CalculateOptimalThreshold()
    sldDitheringAmount.Visible = (cboDither.ListIndex <> 0)
    UpdatePreview
End Sub

'When the auto threshold button is clicked, disable the scroll bar and text box and calculate the optimal value immediately
Private Sub chkAutoThreshold_Click()
    cmdBar.MarkPreviewStatus False
    If chkAutoThreshold.Value Then sltThreshold.Value = CalculateOptimalThreshold()
    sltThreshold.Enabled = Not chkAutoThreshold.Value
    cmdBar.MarkPreviewStatus True
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Color to monochrome", , GetFunctionParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

'When resetting, set the color boxes to black and white, and the dithering combo box to 6 (Stucki)
Private Sub cmdBar_ResetClick()
    
    'Black and white
    csMono(0).Color = RGB(0, 0, 0)
    csMono(1).Color = RGB(255, 255, 255)
    
    'Stucki dithering w/out bleed reduction
    cboDither.ListIndex = 6
    
    'Standard threshold value
    chkAutoThreshold.Value = False
    sltThreshold.Reset
    
End Sub

Private Function GetFunctionParamString() As String
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    With cParams
        .AddParam "threshold", sltThreshold.Value
        .AddParam "dither", cboDither.ListIndex
        .AddParam "ditheramount", sldDitheringAmount.Value
        .AddParam "color1", csMono(0).Color
        .AddParam "color2", csMono(1).Color
        .AddParam "removetransparency", (btsTransparency.ListIndex = 1)
    End With
    GetFunctionParamString = cParams.GetParamString
End Function

Private Sub csMono_ColorChanged(Index As Integer)
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.MarkPreviewStatus False
    
    'Populate the dither combobox
    Palettes.PopulateDitheringDropdown cboDither
    cboDither.ListIndex = 6
    
    btsTransparency.AddItem "do not modify", 0
    btsTransparency.AddItem "remove from image", 1
    btsTransparency.ListIndex = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Calculate the optimal threshold for the current image
Private Function CalculateOptimalThreshold() As Long

    'Create a local array and point it at the pixel data of the image
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    
    EffectPrep.PrepImageData tmpSA, True, pdFxPreview
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
    
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
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Histogram tables
    Dim lLookup(0 To 255) As Long
    Dim pLuminance As Long
    Dim numOfPixels As Long
    
    'Loop through each pixel in the image, tallying values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
            
        'Get the source pixel color values
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
                
        pLuminance = GetLuminance(r, g, b)
        
        'Store this value in the histogram
        lLookup(pLuminance) = lLookup(pLuminance) + 1
        
        'Increment the pixel count
        numOfPixels = numOfPixels + 1
        
    Next y
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    workingDIB.EraseDIB
    Set workingDIB = Nothing
            
    'Divide the number of pixels by two
    numOfPixels = numOfPixels \ 2
                       
    Dim pixelCount As Long
    pixelCount = 0
    x = 0
                    
    'Loop through the histogram table until we have moved past half the pixels in the image
    Do
        pixelCount = pixelCount + lLookup(x)
        x = x + 1
    Loop While pixelCount < numOfPixels
    
    'Make sure our suggestion doesn't exceed the limits allowed by the tool
    If (x > 254) Then x = 220
    
    CalculateOptimalThreshold = x
        
End Function

'Convert an image to black and white (1-bit image)
Public Sub MasterBlackWhiteConversion(ByVal monochromeParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Converting image to two colors..."
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString monochromeParams
    
    Dim cThreshold As Long, ditherMethod As Long, ditherAmount As Single
    Dim lowColor As Long, highColor As Long, removeTransparency As Boolean
    With cParams
        cThreshold = .GetLong("threshold", 127)
        ditherMethod = .GetLong("dither", 6)
        ditherAmount = .GetDouble("ditheramount", 100!)
        lowColor = .GetLong("color1", vbBlack)
        highColor = .GetLong("color2", vbWhite)
        removeTransparency = .GetBool("removetransparency", False)
    End With
    
    ditherAmount = ditherAmount * 0.01
    If (ditherAmount < 0!) Then ditherAmount = 0!
    If (ditherAmount > 1!) Then ditherAmount = 1!
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    
    'If the user wants transparency removed from the image, apply that change prior to monochrome conversion
    Dim alphaAlreadyPremultiplied As Boolean: alphaAlreadyPremultiplied = False
    If (removeTransparency And (curDIBValues.BytesPerPixel = 4)) Then
        EffectPrep.PrepImageData tmpSA, toPreview, dstPic, , , True
        workingDIB.CompositeBackgroundColor 255, 255, 255
        alphaAlreadyPremultiplied = True
    Else
        EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    End If
    
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, i As Long, j As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim xStride As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalX
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Low and high color values
    Dim lowR As Long, lowG As Long, lowB As Long
    Dim highR As Long, highG As Long, highB As Long
    
    lowR = Colors.ExtractRed(lowColor)
    lowG = Colors.ExtractGreen(lowColor)
    lowB = Colors.ExtractBlue(lowColor)
    
    highR = Colors.ExtractRed(highColor)
    highG = Colors.ExtractGreen(highColor)
    highB = Colors.ExtractBlue(highColor)
    
    'Calculating color variables (including luminance)
    Dim r As Long, g As Long, b As Long
    Dim l As Long, newL As Long
    
    Dim ditherTable() As Byte
    Dim xLeft As Long, xRight As Long, yDown As Long
    Dim errorVal As Single
    Dim dDivisor As Single
    
    'Process the image based on the dither method requested
    Select Case ditherMethod
        
        'No dither, so just perform a quick and dirty threshold calculation
        Case 0
    
            For y = initY To finalY
            For x = initX To finalX
                
                xStride = x * qvDepth
                
                'Get the source pixel color values
                b = imageData(xStride, y)
                g = imageData(xStride + 1, y)
                r = imageData(xStride + 2, y)
                
                'Convert those to a luminance value
                l = Colors.GetHQLuminance(r, g, b)
            
                'Check the luminance against the threshold, and set new values accordingly
                If (l >= cThreshold) Then
                    imageData(xStride, y) = highB
                    imageData(xStride + 1, y) = highG
                    imageData(xStride + 2, y) = highR
                Else
                    imageData(xStride, y) = lowB
                    imageData(xStride + 1, y) = lowG
                    imageData(xStride + 2, y) = lowR
                End If
                
            Next x
                If (Not toPreview) Then
                    If (y And progBarCheck) = 0 Then
                        If Interface.UserPressedESC() Then Exit For
                        SetProgBarVal y
                    End If
                End If
            Next y
            
            
        'Ordered dither (Bayer 4x4).  Unfortunately, this routine requires a unique set of code owing to its
        ' specialized implementation. Coefficients derived from http://en.wikipedia.org/wiki/Ordered_dithering
        Case 1
        
            'First, prepare a Bayer dither table
            Palettes.GetDitherTable PDDM_Ordered_Bayer4x4, ditherTable, dDivisor, xLeft, xRight, yDown
            
            'Now loop through the image, using the dither values as our threshold
            For y = initY To finalY
            For x = initX To finalX
            
                xStride = x * qvDepth
                
                'Get the source pixel color values
                b = imageData(xStride, y)
                g = imageData(xStride + 1, y)
                r = imageData(xStride + 2, y)
                
                'Convert those to a luminance value and add the value of the dither table
                l = Colors.GetHQLuminance(r, g, b)
                l = l + (CLng(ditherTable(x And 3, y And 3)) - 127) * ditherAmount
                
                'Check THAT value against the threshold, and set new values accordingly
                If (l >= cThreshold) Then
                    imageData(xStride, y) = highB
                    imageData(xStride + 1, y) = highG
                    imageData(xStride + 2, y) = highR
                Else
                    imageData(xStride, y) = lowB
                    imageData(xStride + 1, y) = lowG
                    imageData(xStride + 2, y) = lowR
                End If
                
            Next x
                If (Not toPreview) Then
                    If (y And progBarCheck) = 0 Then
                        If Interface.UserPressedESC() Then Exit For
                        SetProgBarVal y
                    End If
                End If
            Next y

        'Ordered dither (Bayer 8x8).  Unfortunately, this routine requires a unique set of code owing to its specialized
        ' implementation. Coefficients derived from http://en.wikipedia.org/wiki/Ordered_dithering
        Case 2
        
            'First, prepare a Bayer dither table
            Palettes.GetDitherTable PDDM_Ordered_Bayer8x8, ditherTable, dDivisor, xLeft, xRight, yDown
            
            'Now loop through the image, using the dither values as our threshold
            For y = initY To finalY
            For x = initX To finalX
                
                xStride = x * qvDepth
                
                'Get the source pixel color values
                b = imageData(xStride, y)
                g = imageData(xStride + 1, y)
                r = imageData(xStride + 2, y)
                
                'Convert those to a luminance value and add the value of the dither table
                l = Colors.GetHQLuminance(r, g, b)
                l = l + (CLng(ditherTable(x And 7, y And 7)) - 127) * ditherAmount
                
                'Check THAT value against the threshold, and set new values accordingly
                If (l >= cThreshold) Then
                    imageData(xStride, y) = highB
                    imageData(xStride + 1, y) = highG
                    imageData(xStride + 2, y) = highR
                Else
                    imageData(xStride, y) = lowB
                    imageData(xStride + 1, y) = lowG
                    imageData(xStride + 2, y) = lowR
                End If
                
            Next x
                If (Not toPreview) Then
                    If (y And progBarCheck) = 0 Then
                        If Interface.UserPressedESC() Then Exit For
                        SetProgBarVal y
                    End If
                End If
            Next y
        
        'For all error-diffusion methods, precise dithering table coefficients are retrieved from the
        ' /Modules/Palettes.bas file.  (We do this because other functions also need to retrieve these tables,
        ' e.g. the Effects > Stylize > Palettize menu.)
        
        'Single neighbor.  Simplest form of error-diffusion.
        Case 3
            Palettes.GetDitherTable PDDM_SingleNeighbor, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Genuine Floyd-Steinberg.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 4
            Palettes.GetDitherTable PDDM_FloydSteinberg, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Jarvis, Judice, Ninke.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 5
            Palettes.GetDitherTable PDDM_JarvisJudiceNinke, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Stucki.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 6
            Palettes.GetDitherTable PDDM_Stucki, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Burkes.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 7
            Palettes.GetDitherTable PDDM_Burkes, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Sierra-3.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 8
            Palettes.GetDitherTable PDDM_Sierra3, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Sierra-2.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 9
            Palettes.GetDitherTable PDDM_SierraTwoRow, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Sierra-2-4A.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 10
            Palettes.GetDitherTable PDDM_SierraLite, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Bill Atkinson's original Hyperdither/HyperScan algorithm.  (Note: Bill invented MacPaint, QuickDraw,
        ' and HyperCard.)  This is the dithering algorithm used on the original Apple Macintosh.
        ' Coefficients derived from http://gazs.github.com/canvas-atkinson-dither/
        Case 11
            Palettes.GetDitherTable PDDM_Atkinson, ditherTable, dDivisor, xLeft, xRight, yDown
            
    End Select
    
    'If we have been asked to use a non-ordered dithering method, apply it now
    If (ditherMethod >= PDDM_SingleNeighbor) Then
    
        'First, we need a dithering table the same size as the image.  We make it of Single type to prevent rounding errors.
        ' (This uses a lot of memory, but on modern systems it shouldn't be a problem.)
        Dim dErrors() As Single
        ReDim dErrors(0 To workingDIB.GetDIBWidth, 0 To workingDIB.GetDIBHeight) As Single
        
        If (Not toPreview) Then
            ProgressBars.SetProgBarMax finalY
            progBarCheck = ProgressBars.FindBestProgBarValue()
        End If
        
        Dim xQuick As Long, xQuickInner As Long, yQuick As Long
        
        'Now loop through the image, calculating errors as we go
        For y = initY To finalY
        For x = initX To finalX
            
            xQuick = x * qvDepth
            
            'Get the source pixel color values
            b = imageData(xQuick, y)
            g = imageData(xQuick + 1, y)
            r = imageData(xQuick + 2, y)
            
            'Convert those to a luminance value and add the value of the error at this location
            l = Colors.GetHQLuminance(r, g, b)
            newL = l + dErrors(x, y)
            
            'Check our modified luminance value against the threshold, and set new values accordingly
            If (newL >= cThreshold) Then
                errorVal = newL - 255
                imageData(xQuick, y) = highB
                imageData(xQuick + 1, y) = highG
                imageData(xQuick + 2, y) = highR
            Else
                errorVal = newL
                imageData(xQuick, y) = lowB
                imageData(xQuick + 1, y) = lowG
                imageData(xQuick + 2, y) = lowR
            End If
            
            'If there is an error, spread it
            If (errorVal <> 0) Then
                
                errorVal = errorVal * ditherAmount
                
                'Now, spread that error across the relevant pixels according to the dither table formula
                For i = xLeft To xRight
                For j = 0 To yDown
                
                    'First, ignore already processed pixels
                    If (j = 0) And (i <= 0) Then GoTo NextDitheredPixel
                    
                    'Second, ignore pixels that have a zero in the dither table
                    If ditherTable(i, j) = 0 Then GoTo NextDitheredPixel
                    
                    xQuickInner = x + i
                    yQuick = y + j
                    
                    'Next, ignore target pixels that are off the image boundary
                    If (xQuickInner < initX) Then
                        GoTo NextDitheredPixel
                    ElseIf (xQuickInner > finalX) Then
                        GoTo NextDitheredPixel
                    End If
                    
                    If (yQuick > finalY) Then GoTo NextDitheredPixel
                    
                    'If we've made it all the way here, we are able to actually spread the error to this location
                    dErrors(xQuickInner, yQuick) = dErrors(xQuickInner, yQuick) + (errorVal * (CSng(ditherTable(i, j)) / dDivisor))
                
NextDitheredPixel:     Next j
                Next i
            
            End If
                
        Next x
            If (Not toPreview) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y
                End If
            End If
        Next y
    
    End If
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, alphaAlreadyPremultiplied

End Sub

Private Sub sldDitheringAmount_Change()
    UpdatePreview
End Sub

Private Sub sltThreshold_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then MasterBlackWhiteConversion GetFunctionParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub
