VERSION 5.00
Begin VB.Form FormPosterize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Posterize"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11970
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
   ScaleWidth      =   798
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsAdaptiveColoring 
      Height          =   1020
      Left            =   6000
      TabIndex        =   7
      Top             =   2640
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1799
      Caption         =   "adaptive coloring"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11970
      _ExtentX        =   21114
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
   Begin PhotoDemon.pdSlider sltRed 
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "possible red values"
      Min             =   2
      Max             =   64
      Value           =   6
      DefaultValue    =   6
   End
   Begin PhotoDemon.pdSlider sltGreen 
      Height          =   735
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "possible green values"
      Min             =   2
      Max             =   64
      Value           =   7
      DefaultValue    =   6
   End
   Begin PhotoDemon.pdSlider sltBlue 
      Height          =   735
      Left            =   6000
      TabIndex        =   4
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "possible blue values"
      Min             =   2
      Max             =   64
      Value           =   6
      DefaultValue    =   6
   End
   Begin PhotoDemon.pdSlider sldDitherAmount 
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   4800
      Width           =   5895
      _ExtentX        =   10610
      _ExtentY        =   1296
      Caption         =   "dithering amount"
      Max             =   100
      Value           =   50
      GradientColorRight=   1703935
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdDropDown cboDither 
      Height          =   780
      Left            =   6000
      TabIndex        =   6
      Top             =   3810
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1376
      Caption         =   "dithering"
   End
End
Attribute VB_Name = "FormPosterize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Posterizing Effect Handler
'Copyright 2001-2019 by Tanner Helland
'Created: 4/15/01
'Last updated: 24/May/19
'Last update: overhaul to implement full dithering feature set; also add a bunch of quality and perf improvements
'
'"Posterizing" effect interface.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsAdaptiveColoring_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cboDither_Click()
    UpdateDitherVisibility
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Posterize", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.MarkPreviewStatus False
    
    'Use the palette module to populate our available dithering options
    Palettes.PopulateDitheringDropdown cboDither
    cboDither.ListIndex = 0
    UpdateDitherVisibility
    
    btsAdaptiveColoring.AddItem "off", 0
    btsAdaptiveColoring.AddItem "on", 1
    btsAdaptiveColoring.ListIndex = 0
    
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdateDitherVisibility()
    sldDitherAmount.Visible = (cboDither.ListIndex <> 0)
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.fxPosterize GetLocalParamString(), True, pdFxPreview
End Sub

Public Sub fxPosterize(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    Dim ditherMethod As PD_DITHER_METHOD
    ditherMethod = cParams.GetLong("dithering", 0)
    
    Dim ditherAmount As Single
    ditherAmount = cParams.GetDouble("ditheramount", 100#) * 0.01
    
    'Dithering currently uses an old-school codebase, and thus requires a separate function
    With cParams
        If (ditherMethod = PDDM_None) Then
            ReduceImageColors_BitRGB .GetLong("red"), .GetLong("green"), .GetLong("blue"), .GetBool("matchcolors", True), toPreview, dstPic
        Else
            ReduceImageColors_BitRGB_Dither .GetLong("red"), .GetLong("green"), .GetLong("blue"), ditherMethod, ditherAmount, .GetBool("matchcolors", True), toPreview, dstPic
        End If
    End With
    
End Sub

'Bit RGB color reduction (no error diffusion)
Public Sub ReduceImageColors_BitRGB(ByVal rValue As Byte, ByVal gValue As Byte, ByVal bValue As Byte, Optional ByVal smartColors As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Posterizing image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA2D As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA2D, toPreview, dstPic
    
    Dim pxDepth As Long
    pxDepth = curDIBValues.BytesPerPixel
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left * pxDepth
    initY = curDIBValues.Top
    finalX = curDIBValues.Right * pxDepth
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If smartColors Then ProgressBars.SetProgBarMax finalY * 2 Else ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim mR As Double, mG As Double, mB As Double
    
    'New code for so-called "color matching"
    Dim rLookup() As Long
    Dim gLookup() As Long
    Dim bLookup() As Long
    Dim countLookup() As Long
    
    'Validate input params
    If (rValue > 256) Then rValue = 256
    If (gValue > 256) Then gValue = 256
    If (bValue > 256) Then bValue = 256
    If (rValue < 2) Then rValue = 2
    If (gValue < 2) Then gValue = 2
    If (bValue < 2) Then bValue = 2
    
    rValue = rValue - 1
    gValue = gValue - 1
    bValue = bValue - 1
    
    ReDim rLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim gLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim bLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim countLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    
    'Prepare conversion look-up tables (which will make the actual color reduction much faster)
    mR = (255 / rValue)
    mG = (255 / gValue)
    mB = (255 / bValue)
    
    Dim rQuick(0 To 255) As Byte, gQuick(0 To 255) As Byte, bQuick(0 To 255) As Byte
    For x = 0 To 255
        rQuick(x) = Int((x / mR) + 0.5)
        gQuick(x) = Int((x / mG) + 0.5)
        bQuick(x) = Int((x / mB) + 0.5)
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step pxDepth
        
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Truncate R, G, and B values (posterize-style) into discreet increments.  0.5 is added for rounding purposes.
        newR = rQuick(r)
        newG = gQuick(g)
        newB = bQuick(b)
        
        'If we're doing color matching, place color values into a look-up table
        If smartColors Then
        
            rLookup(newR, newG, newB) = rLookup(newR, newG, newB) + r
            gLookup(newR, newG, newB) = gLookup(newR, newG, newB) + g
            bLookup(newR, newG, newB) = bLookup(newR, newG, newB) + b
            
            'Also, keep track of how many colors fall into this bucket (so we can later determine an average color)
            countLookup(newR, newG, newB) = countLookup(newR, newG, newB) + 1
            
        End If
        
        'Multiply the now-discretely divided R, G, and B values to (0-255) equivalents
        newR = Int(newR * mR + 0.5)
        newG = Int(newG * mG + 0.5)
        newB = Int(newB * mB + 0.5)
        
        'If we are *not* color-matching, assign color values immediately
        If (Not smartColors) Then
            imageData(x) = newB
            imageData(x + 1) = newG
            imageData(x + 2) = newR
        End If
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Color matching requires extra work.  Perform a second loop through the image, replacing values with their
    ' average counterparts.
    If smartColors And (Not g_cancelCurrentAction) Then
    
        'Find average colors based on color counts
        For r = 0 To rValue
        For g = 0 To gValue
        For b = 0 To bValue
            If (countLookup(r, g, b) <> 0) Then
                rLookup(r, g, b) = Int(rLookup(r, g, b) / countLookup(r, g, b) + 0.5)
                gLookup(r, g, b) = Int(gLookup(r, g, b) / countLookup(r, g, b) + 0.5)
                bLookup(r, g, b) = Int(bLookup(r, g, b) / countLookup(r, g, b) + 0.5)
            End If
        Next b
        Next g
        Next r
        
        'Assign average colors back into the picture
        For y = initY To finalY
            workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
        For x = initX To finalX Step pxDepth
            
            newB = bQuick(imageData(x))
            newG = gQuick(imageData(x + 1))
            newR = rQuick(imageData(x + 2))
            
            imageData(x) = bLookup(newR, newG, newB)
            imageData(x + 1) = gLookup(newR, newG, newB)
            imageData(x + 2) = rLookup(newR, newG, newB)
            
        Next x
            If (Not toPreview) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal finalY + y
                End If
            End If
        Next y
        
    End If
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

'Error Diffusion dithering to x# shades of color per component
Public Sub ReduceImageColors_BitRGB_Dither(ByVal rValue As Byte, ByVal gValue As Byte, ByVal bValue As Byte, ByVal ditherType As PD_DITHER_METHOD, ByVal ditherStrength As Single, Optional ByVal smartColors As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Posterizing image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim tmpSA2D As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA2D, toPreview, dstPic
    
    Dim srcPixels1D() As Byte, srcPtr As Long, srcStride As Long
    
    Dim pxSize As Long
    pxSize = workingDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left * pxSize
    initY = curDIBValues.Top
    finalX = curDIBValues.Right * pxSize
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If smartColors Then ProgressBars.SetProgBarMax finalY * 2 Else ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    Dim r As Long, g As Long, b As Long
    Dim i As Long, j As Long
    Dim origR As Long, origG As Long, origB As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim mR As Double, mG As Double, mB As Double
    
    'New code for so-called "color matching"
    Dim rLookup() As Long
    Dim gLookup() As Long
    Dim bLookup() As Long
    Dim countLookup() As Long
    
    'Validate input params
    If (rValue > 256) Then rValue = 256
    If (gValue > 256) Then gValue = 256
    If (bValue > 256) Then bValue = 256
    If (rValue < 2) Then rValue = 2
    If (gValue < 2) Then gValue = 2
    If (bValue < 2) Then bValue = 2
    
    rValue = rValue - 1
    gValue = gValue - 1
    bValue = bValue - 1
    
    ReDim rLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim gLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim bLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim countLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    
    'Prepare conversion look-up tables (which will make the actual color reduction much faster)
    mR = (255 / rValue)
    mG = (255 / gValue)
    mB = (255 / bValue)
    
    Dim rQuick(0 To 255) As Byte, gQuick(0 To 255) As Byte, bQuick(0 To 255) As Byte
    For x = 0 To 255
        rQuick(x) = Int((x / mR) + 0.5)
        gQuick(x) = Int((x / mG) + 0.5)
        bQuick(x) = Int((x / mB) + 0.5)
    Next x
    
    'Validate dither strength
    If (ditherStrength < 0!) Then ditherStrength = 0!
    If (ditherStrength > 1!) Then ditherStrength = 1!
    
    'Prep a dither table that matches the requested setting.  Note that ordered dithers are handled separately.
    Dim ditherTableI() As Byte, ditherDivisor As Single
    Dim xLeft As Long, xRight As Long, yDown As Long
    
    Dim orderedDitherInUse As Boolean
    orderedDitherInUse = (ditherType = PDDM_Ordered_Bayer4x4) Or (ditherType = PDDM_Ordered_Bayer8x8)
    
    If orderedDitherInUse Then
    
        'Ordered dithers are handled specially, because we don't need to track running errors (e.g. no dithering
        ' information is carried to neighboring pixels).  Instead, we simply use the dither tables to adjust our
        ' threshold values on-the-fly.
        Dim ditherRows As Long, ditherColumns As Long
        
        'First, prepare a dithering table
        Palettes.GetDitherTable ditherType, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        
        If (ditherType = PDDM_Ordered_Bayer4x4) Then
            ditherRows = 3
            ditherColumns = 3
        ElseIf (ditherType = PDDM_Ordered_Bayer8x8) Then
            ditherRows = 7
            ditherColumns = 7
        End If
        
        'By default, ordered dither trees use a scale of [0, 255].  This works great for thresholding
        ' against pure black/white, but for color data, it leads to extreme shifts.  Reduce the strength
        ' of the table before continuing.
        For x = 0 To ditherRows
        For y = 0 To ditherColumns
            ditherTableI(x, y) = ditherTableI(x, y) \ 2
        Next y
        Next x
        
        'Apply the finished dither table to the image
        Dim ditherAmt As Long
        
        workingDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            b = srcPixels1D(x)
            g = srcPixels1D(x + 1)
            r = srcPixels1D(x + 2)
            origR = r
            origG = g
            origB = b
            
            'Add dither to each component
            ditherAmt = Int(ditherTableI(Int(x \ 4) And ditherRows, y And ditherColumns)) - 63
            ditherAmt = ditherAmt * ditherStrength
            
            r = r + ditherAmt
            If (r > 255) Then
                r = 255
            ElseIf (r < 0) Then
                r = 0
            End If
            
            g = g + ditherAmt
            If (g > 255) Then
                g = 255
            ElseIf (g < 0) Then
                g = 0
            End If
            
            b = b + ditherAmt
            If (b > 255) Then
                b = 255
            ElseIf (b < 0) Then
                b = 0
            End If
            
            'Posterize
            newR = rQuick(r)
            newG = gQuick(g)
            newB = bQuick(b)
            
            'If we're doing color matching, place color values into a look-up table
            If smartColors Then
            
                rLookup(newR, newG, newB) = rLookup(newR, newG, newB) + origR
                gLookup(newR, newG, newB) = gLookup(newR, newG, newB) + origG
                bLookup(newR, newG, newB) = bLookup(newR, newG, newB) + origB
                
                'Also, keep track of how many colors fall into this bucket (so we can later determine an average color)
                countLookup(newR, newG, newB) = countLookup(newR, newG, newB) + 1
                
            End If
                
            'Multiply the now-discretely divided R, G, and B values to (0-255) equivalents
            newR = Int(newR * mR + 0.5)
            newG = Int(newG * mG + 0.5)
            newB = Int(newB * mB + 0.5)
            
            srcPixels1D(x) = newB
            srcPixels1D(x + 1) = newG
            srcPixels1D(x + 2) = newR
            
        Next x
            If (Not toPreview) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y
                End If
            End If
        Next y
        
        workingDIB.UnwrapArrayFromDIB srcPixels1D
    
    'All error-diffusion dither methods are handled similarly
    Else
        
        Dim rError As Long, gError As Long, bError As Long
        Dim errorMult As Single
        
        'Retrieve a hard-coded dithering table matching the requested dither type
        Palettes.GetDitherTable ditherType, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        If (ditherDivisor <> 0!) Then ditherDivisor = 1! / ditherDivisor
        
        'Next, build an error tracking array.  Some diffusion methods require three rows worth of others;
        ' others require two.  Note that errors must be tracked separately for each color component.
        Dim xWidth As Long
        xWidth = workingDIB.GetDIBWidth - 1
        
        Dim rErrors() As Single, gErrors() As Single, bErrors() As Single
        ReDim rErrors(0 To xWidth, 0 To yDown) As Single
        ReDim gErrors(0 To xWidth, 0 To yDown) As Single
        ReDim bErrors(0 To xWidth, 0 To yDown) As Single
        
        Dim xNonStride As Long, xQuickInner As Long
        'Dim newR As Long, newG As Long, newB As Long
        
        workingDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        'Start calculating pixels.
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            b = srcPixels1D(x)
            g = srcPixels1D(x + 1)
            r = srcPixels1D(x + 2)
            origR = r
            origG = g
            origB = b
            
            'Add our running errors to the original colors
            xNonStride = x \ 4
            r = origR + rErrors(xNonStride, 0)
            g = origG + gErrors(xNonStride, 0)
            b = origB + bErrors(xNonStride, 0)
            
            If (r > 255) Then
                r = 255
            ElseIf (r < 0) Then
                r = 0
            End If
            
            If (g > 255) Then
                g = 255
            ElseIf (g < 0) Then
                g = 0
            End If
            
            If (b > 255) Then
                b = 255
            ElseIf (b < 0) Then
                b = 0
            End If
            
            'Posterize
            newR = rQuick(r)
            newG = gQuick(g)
            newB = bQuick(b)
            
            'If we're doing color matching, place color values into a look-up table
            If smartColors Then
            
                rLookup(newR, newG, newB) = rLookup(newR, newG, newB) + origR
                gLookup(newR, newG, newB) = gLookup(newR, newG, newB) + origG
                bLookup(newR, newG, newB) = bLookup(newR, newG, newB) + origB
                
                'Also, keep track of how many colors fall into this bucket (so we can later determine an average color)
                countLookup(newR, newG, newB) = countLookup(newR, newG, newB) + 1
                
            End If
            
            'Multiply the now-discretely divided R, G, and B values to (0-255) equivalents
            newR = Int(newR * mR + 0.5)
            newG = Int(newG * mG + 0.5)
            newB = Int(newB * mB + 0.5)
            
            srcPixels1D(x) = newB
            srcPixels1D(x + 1) = newG
            srcPixels1D(x + 2) = newR
            
            'Calculate new errors
            rError = r - newR
            gError = g - newG
            bError = b - newB
            
            'Reduce color bleed, if specified
            rError = rError * ditherStrength
            gError = gError * ditherStrength
            bError = bError * ditherStrength
            
            'Spread any remaining error to neighboring pixels, using the precalculated dither table as our guide
            For i = xLeft To xRight
            For j = 0 To yDown
                
                If (ditherTableI(i, j) <> 0) Then
                    
                    xQuickInner = xNonStride + i
                    
                    'Next, ignore target pixels that are off the image boundary
                    If (xQuickInner >= initX) Then
                        If (xQuickInner < xWidth) Then
                        
                            'If we've made it all the way here, we are able to actually spread the error to this location
                            errorMult = CSng(ditherTableI(i, j)) * ditherDivisor
                            rErrors(xQuickInner, j) = rErrors(xQuickInner, j) + (rError * errorMult)
                            gErrors(xQuickInner, j) = gErrors(xQuickInner, j) + (gError * errorMult)
                            bErrors(xQuickInner, j) = bErrors(xQuickInner, j) + (bError * errorMult)
                            
                        End If
                    End If
                    
                End If
                
            Next j
            Next i
            
        Next x
        
            'When moving to the next line, we need to "shift" all accumulated errors upward.
            ' (Basically, what was previously the "next" line, is now the "current" line.
            ' The last line of errors must also be zeroed-out.
            If (yDown > 0) Then
            
                CopyMemory ByVal VarPtr(rErrors(0, 0)), ByVal VarPtr(rErrors(0, 1)), (xWidth + 1) * 4
                CopyMemory ByVal VarPtr(gErrors(0, 0)), ByVal VarPtr(gErrors(0, 1)), (xWidth + 1) * 4
                CopyMemory ByVal VarPtr(bErrors(0, 0)), ByVal VarPtr(bErrors(0, 1)), (xWidth + 1) * 4
                
                If (yDown = 1) Then
                    FillMemory VarPtr(rErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 1)), (xWidth + 1) * 4, 0
                Else
                    CopyMemory ByVal VarPtr(rErrors(0, 1)), ByVal VarPtr(rErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemory ByVal VarPtr(gErrors(0, 1)), ByVal VarPtr(gErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemory ByVal VarPtr(bErrors(0, 1)), ByVal VarPtr(bErrors(0, 2)), (xWidth + 1) * 4
                    
                    FillMemory VarPtr(rErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 2)), (xWidth + 1) * 4, 0
                End If
                
            Else
                FillMemory VarPtr(rErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(gErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(bErrors(0, 0)), (xWidth + 1) * 4, 0
            End If
            
            'Update the progress bar, as necessary
            If (Not toPreview) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y
                End If
            End If
            
        Next y
        
        workingDIB.UnwrapArrayFromDIB srcPixels1D
    
    End If
    
    'Color matching requires extra work.  Perform a second loop through the image, replacing values with their
    ' average counterparts.
    If smartColors And (Not g_cancelCurrentAction) Then
    
        'Find average colors based on color counts
        For r = 0 To rValue
        For g = 0 To gValue
        For b = 0 To bValue
            If (countLookup(r, g, b) <> 0) Then
                rLookup(r, g, b) = Int(rLookup(r, g, b) / countLookup(r, g, b) + 0.5)
                gLookup(r, g, b) = Int(gLookup(r, g, b) / countLookup(r, g, b) + 0.5)
                bLookup(r, g, b) = Int(bLookup(r, g, b) / countLookup(r, g, b) + 0.5)
            End If
        Next b
        Next g
        Next r
        
        'Assign average colors back into the picture
        For y = initY To finalY
            workingDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, y
        For x = initX To finalX Step pxSize
            
            newB = bQuick(srcPixels1D(x))
            newG = gQuick(srcPixels1D(x + 1))
            newR = rQuick(srcPixels1D(x + 2))
            
            srcPixels1D(x) = bLookup(newR, newG, newB)
            srcPixels1D(x + 1) = gLookup(newR, newG, newB)
            srcPixels1D(x + 2) = rLookup(newR, newG, newB)
            
        Next x
            If (Not toPreview) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal finalY + y
                End If
            End If
        Next y
        
        workingDIB.UnwrapArrayFromDIB srcPixels1D
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub sldDitherAmount_Change()
    UpdatePreview
End Sub

Private Sub sltBlue_Change()
    UpdatePreview
End Sub

Private Sub sltGreen_Change()
    UpdatePreview
End Sub

Private Sub sltRed_Change()
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
        
        .AddParam "red", sltRed.Value
        .AddParam "green", sltGreen.Value
        .AddParam "blue", sltBlue.Value
        
        .AddParam "matchcolors", (btsAdaptiveColoring.ListIndex = 1)
        
        .AddParam "dithering", cboDither.ListIndex
        .AddParam "ditheramount", sldDitherAmount.Value
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
