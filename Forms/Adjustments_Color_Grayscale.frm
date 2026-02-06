VERSION 5.00
Begin VB.Form FormGrayscale 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Black and white"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11895
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
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   Visible         =   0   'False
   Begin PhotoDemon.pdSlider sldDitherAmount 
      Height          =   735
      Left            =   6000
      TabIndex        =   7
      Top             =   4080
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1296
      Caption         =   "dithering amount"
      Max             =   100
      Value           =   100
      DefaultValue    =   100
   End
   Begin PhotoDemon.pdDropDown cboDithering 
      Height          =   735
      Left            =   6000
      TabIndex        =   1
      Top             =   3000
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1296
      Caption         =   "dithering"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5805
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltShades 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2040
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1244
      Caption         =   "number of gray shades"
      Min             =   2
      Max             =   256
      Value           =   256
      NotchPosition   =   2
      NotchValueCustom=   256
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdDropDown cboMethod 
      Height          =   735
      Left            =   6000
      TabIndex        =   4
      Top             =   360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1296
      Caption         =   "style"
   End
   Begin PhotoDemon.pdButtonStrip btsChannel 
      Height          =   630
      Left            =   6120
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1111
   End
   Begin PhotoDemon.pdButtonStrip btsDecompose 
      Height          =   630
      Left            =   6120
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1111
   End
End
Attribute VB_Name = "FormGrayscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Grayscale Conversion Handler
'Copyright 2002-2026 by Tanner Helland
'Created: 1/12/02
'Last updated: 20/April/20
'Last update: condense conversions into a single function; perf optimize that function;
'             improve UI reflow when options are toggled
'
'Updated version of the grayscale handler; utilizes five different methods (average, ITU,
' desaturate, max/min decomposition, single color channel) with the option for variable
' # of gray shades with/without dithering for all available methods. A comprehensive dithering
' list is also available for all methods, should the user desire it, including adjustable
' strength.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Enum PD_GrayscaleTechnique
    GT_Fast = 0
    GT_ITU = 1
    GT_Desaturate = 2
    GT_Decompose = 3
    GT_Channel = 4
End Enum

#If False Then
    Private Const GT_Fast = 0, GT_ITU = 1, GT_Desaturate = 2, GT_Decompose = 3, GT_Channel = 4
#End If

'Preview the current grayscale conversion technique
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then GrayscaleConvert_Central GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub btsChannel_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsDecompose_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cboDithering_Click()
    UpdateVisibleControls
    UpdatePreview
End Sub

Private Sub cboMethod_Click()
    UpdateVisibleControls
    UpdatePreview
End Sub

'Certain algorithms require additional user input.  This routine enables/disables the controls associated with a given algorithm.
Private Sub UpdateVisibleControls()

    Dim yOffset As Long
    yOffset = cboMethod.GetTop + cboMethod.GetHeight + Interface.FixDPI(8)
    
    If (cboMethod.ListIndex = GT_Decompose) Then
        btsDecompose.SetTop yOffset
        yOffset = yOffset + btsDecompose.GetHeight + Interface.FixDPI(8)
        btsDecompose.Visible = True
    Else
        btsDecompose.Visible = False
    End If
    
    If (cboMethod.ListIndex = GT_Channel) Then
        btsChannel.SetTop yOffset
        yOffset = yOffset + btsChannel.GetHeight + Interface.FixDPI(8)
        btsChannel.Visible = True
    Else
        btsChannel.Visible = False
    End If
    
    yOffset = yOffset + Interface.FixDPI(4)
    sltShades.SetTop yOffset
    yOffset = yOffset + sltShades.GetHeight + Interface.FixDPI(8)
    
    If (sltShades.Value <> 256) Then
        cboDithering.SetTop yOffset
        yOffset = yOffset + cboDithering.GetHeight + Interface.FixDPI(12)
        cboDithering.Visible = True
        
        If (cboDithering.ListIndex <> 0) Then
            sldDitherAmount.SetTop yOffset
            sldDitherAmount.Visible = True
        Else
            sldDitherAmount.Visible = False
        End If
        
    Else
        cboDithering.Visible = False
        sldDitherAmount.Visible = False
    End If
    
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Black and white", False, GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdateVisibleControls
    UpdatePreview
End Sub

'Recommend ITU grayscale correction by default, and max shades without dithering
Private Sub cmdBar_ResetClick()
    cboMethod.ListIndex = 1
    cboDithering.ListIndex = 6
    sltShades.Value = 256
End Sub

'All different grayscale (black and white) routines are handled by this single function.  As of 16 Feb '14, grayscale operations
' are divided into four params: type of transform, optional params for transform (if any), number of shades to use, and
' dithering options (if any).  This should allow the user to mix and match the various options at their leisure.
Public Sub GrayscaleConvert_Central(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Converting image to black and white..."
    
    Dim grayscaleMethod As PD_GrayscaleTechnique, numOfShades As Long, ditheringOptions As PD_DITHER_METHOD, ditherAmount As Single
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    With cParams
        grayscaleMethod = .GetLong("method", GT_ITU)
        numOfShades = .GetLong("shades", 256)
        ditheringOptions = .GetLong("dithering", 0)
        ditherAmount = .GetDouble("ditheramount", 100!) * 0.01!
    End With
    
    If (ditherAmount < 0!) Then ditherAmount = 0!
    If (ditherAmount > 1!) Then ditherAmount = 1!
    
    'Create a working copy of the relevant pixel data (with all selection transforms applied)
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Based on the options the user has provided, figure out a maximum progress bar value.  This changes depending on:
    ' - If the user wants shade reduction (as this requires another pass over the image)
    ' - If the user wants dithering (also requires a second pass)
    Dim progBarMax As Long
    progBarMax = workingDIB.GetDIBHeight
    If (numOfShades < 256) Then progBarMax = progBarMax + workingDIB.GetDIBHeight
    
    'Some conversion methods support extra parameters; retrieve those now
    Dim bonusParams As Long
    If (grayscaleMethod = GT_Decompose) Then bonusParams = cParams.GetLong("decomposemode", 0)
    If (grayscaleMethod = GT_Channel) Then bonusParams = cParams.GetLong("channelmode", 1)
    
    'Convert to grayscale
    Dim userCanceled As Long
    userCanceled = ConvertToGrayscale(grayscaleMethod, bonusParams, workingDIB, toPreview, progBarMax)
    
    'We now apply the user's choice of shade reduction and/or dithering.
    If (numOfShades < 256) And (userCanceled <> 0) Then
        
        Select Case ditheringOptions
        
            Case PDDM_None
                fGrayscaleCustom numOfShades, workingDIB, toPreview, progBarMax, workingDIB.GetDIBHeight
            
            'If dithering is active, we can simply build a grayscale palette, then ask the central Palette engine to
            ' do the work for us.
            Case Else
                fGrayscaleCustomDither numOfShades, ditheringOptions, ditherAmount, workingDIB, toPreview, progBarMax, workingDIB.GetDIBHeight
            
        End Select
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

'Reduce to X # gray shades
Private Function fGrayscaleCustom(ByVal numOfShades As Long, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Point an array at the source DIB's image data
    Dim imageData() As Byte, srcSA As SafeArray1D
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = (srcDIB.GetDIBWidth - 1) * 4
    finalY = srcDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color variables
    Dim r As Long, g As Long, b As Long, grayVal As Long
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Double
    conversionFactor = (255# / (numOfShades - 1))
    
    'Build a look-up table for our custom grayscale conversion results
    Dim gLookup(0 To 255) As Byte
    
    For x = 0 To 255
        grayVal = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
        If (grayVal > 255) Then grayVal = 255
        gLookup(x) = CByte(grayVal)
    Next x
    
    'Build another look-up table for our initial grayscale index calculation
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        srcDIB.WrapArrayAroundScanline imageData, srcSA, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        grayVal = gLookup(grayLookUp(r + g + b))
        
        'Assign all color channels the new gray value
        imageData(x) = grayVal
        imageData(x + 1) = grayVal
        imageData(x + 2) = grayVal
        
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    srcDIB.UnwrapArrayFromDIB imageData
    
    If g_cancelCurrentAction Then fGrayscaleCustom = 0 Else fGrayscaleCustom = 1
    
End Function

'Reduce to X # gray shades (dithered)
Private Function fGrayscaleCustomDither(ByVal numOfShades As Long, ByVal ditherMethod As Long, ByVal ditherAmount As Single, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    Dim srcPixels() As Byte, tmpSA As SafeArray2D
    dstDIB.WrapArrayAroundDIB srcPixels, tmpSA
    
    Dim srcPixels1D() As Byte, tmpSA1D As SafeArray1D, srcPtr As Long, srcStride As Long
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color variables (we don't need full RGB calculations - just one channel will do)
    Dim g As Long, newG As Long
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Double
    conversionFactor = (255# / (numOfShades - 1))
    
    'Build a look-up table for our custom grayscale conversion results
    Dim gLookup(0 To 255) As Long
    For x = 0 To 255
        newG = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
        If (newG > 255) Then newG = 255
        gLookup(x) = newG
    Next x
    
    'Prep a dither table that matches the requested setting.  Note that ordered dithers are handled separately.
    Dim ditherTableI() As Byte, ditherDivisor As Single
    Dim xLeft As Long, xRight As Long, yDown As Long
    
    Dim orderedDitherInUse As Boolean
    orderedDitherInUse = (ditherMethod = PDDM_Ordered_Bayer4x4) Or (ditherMethod = PDDM_Ordered_Bayer8x8)
    
    If orderedDitherInUse Then
    
        'Ordered dithers are handled specially, because we don't need to track running errors (e.g. no dithering
        ' information is carried to neighboring pixels).  Instead, we simply use the dither tables to adjust our
        ' threshold values on-the-fly.
        Dim ditherRows As Long, ditherColumns As Long
        
        'First, prepare a dithering table
        Palettes.GetDitherTable ditherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        
        If (ditherMethod = PDDM_Ordered_Bayer4x4) Then
            ditherRows = 3
            ditherColumns = 3
        ElseIf (ditherMethod = PDDM_Ordered_Bayer8x8) Then
            ditherRows = 7
            ditherColumns = 7
        End If
        
        'By default, ordered dither trees use a scale of [0, 255].  This works great for thresholding
        ' against pure black/white, but for other gradients, it leads to extreme shifts.  Reduce the
        ' strength of the table before continuing.
        For x = 0 To ditherRows
        For y = 0 To ditherColumns
            ditherTableI(x, y) = ditherTableI(x, y) \ 2
        Next y
        Next x
        
        'Apply the finished dither table to the image
        Dim ditherAmt As Long
        
        dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            'Get the source pixel color values.  Because we know the image we're handed is already going to be grayscale,
            ' we can shortcut this calculation by only grabbing one channel.
            g = srcPixels1D(x)
            
            'Add dither
            ditherAmt = Int(ditherTableI(Int(x \ 4) And ditherRows, y And ditherColumns)) - 64
            ditherAmt = ditherAmt * ditherAmount
            
            'Convert those to a luminance value and add the value of the error at this location
            newG = g + ditherAmt
            
            'Convert that to a lookup-table-safe luminance (e.g. 0-255)
            If (newG < 0) Then
                newG = 0
            ElseIf (newG > 255) Then
                newG = 255
            End If
            
            'Write the new luminance value out to the image array
            newG = gLookup(newG)
            srcPixels1D(x) = newG
            srcPixels1D(x + 1) = newG
            srcPixels1D(x + 2) = newG
            
        Next x
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If
        Next y
        
        dstDIB.UnwrapArrayFromDIB srcPixels1D
    
    'All error-diffusion dither methods are handled similarly
    Else
        
        Dim gError As Long, finalG As Long
        Dim errorMult As Single
        
        'Retrieve a hard-coded dithering table matching the requested dither type
        Palettes.GetDitherTable ditherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        If (ditherDivisor <> 0!) Then ditherDivisor = 1! / ditherDivisor
        
        'Next, build an error tracking array.  Some diffusion methods require three rows worth of others;
        ' others require two.  Note that errors must be tracked separately for each color component.
        Dim xWidth As Long
        xWidth = workingDIB.GetDIBWidth - 1
        Dim gErrors() As Single
        ReDim gErrors(0 To xWidth, 0 To yDown) As Single
        
        Dim xNonStride As Long, xQuickInner As Long
        
        dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        'Start calculating pixels.
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            g = srcPixels1D(x)
            
            'Add our running errors to the original colors
            xNonStride = x \ 4
            newG = g + gErrors(xNonStride, 0)
            
            If (newG > 255) Then
                newG = 255
            ElseIf (newG < 0) Then
                newG = 0
            End If
            
            'Calculate the matching color
            finalG = gLookup(newG)
            
            'Apply the closest discovered color to this pixel.
            srcPixels1D(x) = finalG
            srcPixels1D(x + 1) = finalG
            srcPixels1D(x + 2) = finalG
            
            'Calculate new error
            gError = newG - finalG
            
            'Reduce color bleed, if specified
            gError = gError * ditherAmount
            
            'Spread any remaining error to neighboring pixels, using the precalculated dither table as our guide
            Dim i As Long, j As Long
            For i = xLeft To xRight
            For j = 0 To yDown
                
                If (ditherTableI(i, j) <> 0) Then
                    
                    xQuickInner = xNonStride + i
                    
                    'Next, ignore target pixels that are off the image boundary
                    If (xQuickInner >= initX) Then
                        If (xQuickInner < xWidth) Then
                        
                            'If we've made it all the way here, we are able to actually spread the error to this location
                            errorMult = ditherTableI(i, j) * ditherDivisor
                            gErrors(xQuickInner, j) = gErrors(xQuickInner, j) + (gError * errorMult)
                            
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
            
                CopyMemoryStrict VarPtr(gErrors(0, 0)), VarPtr(gErrors(0, 1)), (xWidth + 1) * 4
                
                If (yDown = 1) Then
                    FillMemory VarPtr(gErrors(0, 1)), (xWidth + 1) * 4, 0
                Else
                    CopyMemoryStrict VarPtr(gErrors(0, 1)), VarPtr(gErrors(0, 2)), (xWidth + 1) * 4
                    FillMemory VarPtr(gErrors(0, 2)), (xWidth + 1) * 4, 0
                End If
                
            Else
                FillMemory VarPtr(gErrors(0, 0)), (xWidth + 1) * 4, 0
            End If
            
            'Update the progress bar, as necessary
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If
            
        Next y
        
        dstDIB.UnwrapArrayFromDIB srcPixels1D
    
    End If
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    If g_cancelCurrentAction Then fGrayscaleCustomDither = 0 Else fGrayscaleCustomDither = 1
    
End Function

'Reduce to gray in a more human-eye friendly manner
Private Function ConvertToGrayscale(ByVal grayscaleMethod As PD_GrayscaleTechnique, ByVal bonusParams As Long, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Point an array at the source DIB's image data
    Dim imageData() As Byte, srcSA As SafeArray1D
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = (srcDIB.GetDIBWidth - 1) * 4
    finalY = srcDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
    
    'Some conversion techniques work well with prebuilt LUTs
    Dim grayLookUp(0 To 765) As Byte
    If (grayscaleMethod = GT_Fast) Then
        For x = 0 To 765
            grayLookUp(x) = x \ 3
        Next x
    End If
    
        
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        srcDIB.WrapArrayAroundScanline imageData, srcSA, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        Select Case grayscaleMethod
            
            '(r + g + b) / 3
            Case GT_Fast
                grayVal = grayLookUp(r + g + b)
            
            'original ITU-R recommended formula (BT.709, specifically)
            Case GT_ITU
                grayVal = (218 * r + 732 * g + 74 * b) \ 1024
            
            '(max + min) / 2 - this is HSL with saturation forced to zero
            Case GT_Desaturate
                grayVal = Colors.GetLuminance(r, g, b)
            
            'Max(r, g, b) or Min(r, g, b)
            Case GT_Decompose
                If (bonusParams = 0) Then grayVal = Min3Int(r, g, b) Else grayVal = Max3Int(r, g, b)
            
            'Just use r, g, or b as-is
            Case GT_Channel
                Select Case bonusParams
                    Case 0
                        grayVal = r
                    Case 1
                        grayVal = g
                    Case 2
                        grayVal = b
                End Select
            
        End Select
        
        'Assign the calculated gray value to each color channel
        imageData(x) = grayVal
        imageData(x + 1) = grayVal
        imageData(x + 2) = grayVal
        
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    srcDIB.UnwrapArrayFromDIB imageData
    
    If g_cancelCurrentAction Then ConvertToGrayscale = 0 Else ConvertToGrayscale = 1
    
End Function

Private Sub Form_Load()
    
    'Suspend previews while we get the form set up
    cmdBar.SetPreviewStatus False
    
    'Set up the grayscale options combo box
    cboMethod.SetAutomaticRedraws False
    cboMethod.Clear
    cboMethod.AddItem "Fastest Calculation (average value)", 0
    cboMethod.AddItem "Highest Quality (ITU Standard)", 1
    cboMethod.AddItem "Desaturate", 2
    cboMethod.AddItem "Decompose", 3
    cboMethod.AddItem "Single color channel", 4
    cboMethod.SetAutomaticRedraws True
    cboMethod.ListIndex = 1
    
    'Populate the dither dropdown
    Palettes.PopulateDitheringDropdown cboDithering
    cboDithering.ListIndex = 6
    
    'Populate any other per-method controls
    btsDecompose.AddItem "minimum", 0
    btsDecompose.AddItem "maximum", 1
    
    btsChannel.AddItem "red", 0
    btsChannel.AddItem "green", 1
    btsChannel.AddItem "blue", 2
    
    'Make sure the correct options subpanel is set
    UpdateVisibleControls
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    
    'Draw the initial preview
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sldDitherAmount_Change()
    UpdatePreview
End Sub

Private Sub sltShades_Change()
    UpdateVisibleControls
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
        
        'Three parameters are always relevant, regardless of the current grayscale algorithm
        .AddParam "method", cboMethod.ListIndex
        .AddParam "shades", sltShades.Value
        .AddParam "dithering", cboDithering.ListIndex
        .AddParam "ditheramount", sldDitherAmount.Value
        
        'All following parameters are relevant to only certain grayscale modes.
        .AddParam "decomposemode", btsDecompose.ListIndex
        .AddParam "channelmode", btsChannel.ListIndex
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
