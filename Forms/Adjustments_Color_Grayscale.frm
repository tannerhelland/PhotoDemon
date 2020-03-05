VERSION 5.00
Begin VB.Form FormGrayscale 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Black and white"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11895
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
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   ShowInTaskbar   =   0   'False
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
   Begin PhotoDemon.pdButtonStrip btsDecompose 
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   1260
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   873
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
      Align           =   2  'Align Bottom
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
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1270
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
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   1260
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   873
   End
End
Attribute VB_Name = "FormGrayscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Grayscale Conversion Handler
'Copyright 2002-2020 by Tanner Helland
'Created: 1/12/02
'Last updated: 02/April/18
'Last update: add "single neighbor" as a dithering option
'
'Updated version of the grayscale handler; utilizes five different methods (average, ISU, desaturate, max/min decomposition,
' single color channel) with the option for variable # of gray shades with/without dithering for all available methods. A
' comprehensive dithering list is also available for all methods, should the user desire it.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
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
    If cmdBar.PreviewsAllowed Then MasterGrayscaleFunction GetLocalParamString(), True, pdFxPreview
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
    btsDecompose.Visible = (cboMethod.ListIndex = GT_Decompose)
    btsChannel.Visible = (cboMethod.ListIndex = GT_Channel)
    cboDithering.Visible = (sltShades.Value <> 256)
    sldDitherAmount.Visible = (sltShades.Value <> 256) And (cboDithering.ListIndex <> 0)
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
Public Sub MasterGrayscaleFunction(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Converting image to black and white..."
    
    Dim grayscaleMethod As PD_GrayscaleTechnique, numOfShades As Long, ditheringOptions As PD_DITHER_METHOD, ditherAmount As Single
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    With cParams
        
        'Three parameters are always relevant, regardless of the current grayscale algorithm
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
    ' - If the user wants dithering (as the second pass will be done horizontally instead of vertically)
    Dim progBarMax As Long
    If (numOfShades < 256) Then
        progBarMax = workingDIB.GetDIBWidth + workingDIB.GetDIBHeight
    Else
        progBarMax = workingDIB.GetDIBWidth
    End If
    
    Dim userCanceled As Long
    
    'Different grayscale conversion methods call different individual subs
    Select Case grayscaleMethod
        
        Case GT_Fast
            userCanceled = MenuGrayscaleAverage(workingDIB, toPreview, progBarMax)
            
        Case GT_ITU
            userCanceled = MenuGrayscale(workingDIB, toPreview, progBarMax)
            
        Case GT_Desaturate
            userCanceled = MenuDesaturate(workingDIB, toPreview, progBarMax)
            
        Case GT_Decompose
            userCanceled = MenuDecompose(cParams.GetLong("decomposemode", 0), workingDIB, toPreview, progBarMax)
            
        Case GT_Channel
            userCanceled = MenuGrayscaleSingleChannel(cParams.GetLong("channelmode", 1), workingDIB, toPreview, progBarMax)
            
    End Select
    
    'We now apply the user's choice of shade reduction and/or dithering.
    If (numOfShades < 256) And (userCanceled <> 0) Then
        
        Select Case ditheringOptions
        
            Case PDDM_None
                fGrayscaleCustom numOfShades, workingDIB, toPreview, progBarMax, workingDIB.GetDIBWidth
            
            'If dithering is active, we can simply build a grayscale palette, then ask the central Palette engine to
            ' do the work for us.
            Case Else
                fGrayscaleCustomDither numOfShades, ditheringOptions, ditherAmount, workingDIB, toPreview, progBarMax, workingDIB.GetDIBWidth
            
        End Select
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

'Reduce to X # gray shades
Public Function fGrayscaleCustom(ByVal numOfShades As Long, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Point an array at the source DIB's image data
    Dim imageData() As Byte, srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
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
    For x = initX To finalX
        
        quickVal = x * qvDepth
        
        'Get the source pixel color values
        b = imageData(quickVal, y)
        g = imageData(quickVal + 1, y)
        r = imageData(quickVal + 2, y)
        
        grayVal = grayLookUp(r + g + b)
        
        'Assign all color channels the new gray value
        grayVal = gLookup(grayVal)
        imageData(quickVal, y) = grayVal
        imageData(quickVal + 1, y) = grayVal
        imageData(quickVal + 2, y) = grayVal
        
    Next x
        If Not suppressMessages Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then fGrayscaleCustom = 0 Else fGrayscaleCustom = 1
    
End Function

'Reduce to X # gray shades (dithered)
Public Function fGrayscaleCustomDither(ByVal numOfShades As Long, ByVal ditherMethod As Long, ByVal ditherAmount As Single, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

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
                            errorMult = CSng(ditherTableI(i, j)) * ditherDivisor
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
            
                CopyMemory ByVal VarPtr(gErrors(0, 0)), ByVal VarPtr(gErrors(0, 1)), (xWidth + 1) * 4
                
                If (yDown = 1) Then
                    FillMemory VarPtr(gErrors(0, 1)), (xWidth + 1) * 4, 0
                Else
                    CopyMemory ByVal VarPtr(gErrors(0, 1)), ByVal VarPtr(gErrors(0, 2)), (xWidth + 1) * 4
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

'Reduce to gray via (r+g+b)/3
Public Function MenuGrayscaleAverage(ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Point an array at the source DIB's image data
    Dim imageData() As Byte, srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
    
    'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        b = imageData(quickVal, y)
        g = imageData(quickVal + 1, y)
        r = imageData(quickVal + 2, y)
        
        'Calculate the gray value using the look-up table
        grayVal = grayLookUp(r + g + b)
        
        'Assign that gray value to each color channel
        imageData(quickVal, y) = grayVal
        imageData(quickVal + 1, y) = grayVal
        imageData(quickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then MenuGrayscaleAverage = 0 Else MenuGrayscaleAverage = 1
    
End Function

'Reduce to gray in a more human-eye friendly manner
Public Function MenuGrayscale(ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Point an array at the source DIB's image data
    Dim imageData() As Byte, srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        b = imageData(quickVal, y)
        g = imageData(quickVal + 1, y)
        r = imageData(quickVal + 2, y)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        grayVal = (213 * r + 715 * g + 72 * b) \ 1000
        If (grayVal > 255) Then grayVal = 255
        
        'Assign that gray value to each color channel
        imageData(quickVal, y) = grayVal
        imageData(quickVal + 1, y) = grayVal
        imageData(quickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then MenuGrayscale = 0 Else MenuGrayscale = 1
    
End Function

'Reduce to gray via HSL -> convert S to 0
Public Function MenuDesaturate(ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
        
    'Point an array at the source DIB's image data
    Dim imageData() As Byte, srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
       
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        b = imageData(quickVal, y)
        g = imageData(quickVal + 1, y)
        r = imageData(quickVal + 2, y)
        
        'Calculate a grayscale value by using a short-hand RGB <-> HSL conversion
        grayVal = CByte(GetLuminance(r, g, b))
        
        'Assign that gray value to each color channel
        imageData(quickVal, y) = grayVal
        imageData(quickVal + 1, y) = grayVal
        imageData(quickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then MenuDesaturate = 0 Else MenuDesaturate = 1
    
End Function

'Reduce to gray by selecting the minimum (maxOrMin = 0) or maximum (maxOrMin = 1) color in each pixel
Public Function MenuDecompose(ByVal maxOrMin As Long, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Point an array at the source DIB's image data
    Dim imageData() As Byte, srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        b = imageData(quickVal, y)
        g = imageData(quickVal + 1, y)
        r = imageData(quickVal + 2, y)
        
        'Find the highest or lowest of the RGB values
        If maxOrMin = 0 Then grayVal = CByte(Min3Int(r, g, b)) Else grayVal = CByte(Max3Int(r, g, b))
        
        'Assign that gray value to each color channel
        imageData(quickVal, y) = grayVal
        imageData(quickVal + 1, y) = grayVal
        imageData(quickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then MenuDecompose = 0 Else MenuDecompose = 1
    
End Function

'Reduce to gray by selecting a single color channel (represeted by cChannel: 0 = Red, 1 = Green, 2 = Blue)
Public Function MenuGrayscaleSingleChannel(ByVal cChannel As Long, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Point an array at the source DIB's image data
    Dim imageData() As Byte, srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        b = imageData(quickVal, y)
        g = imageData(quickVal + 1, y)
        r = imageData(quickVal + 2, y)
        
        'Assign the gray value to a single color channel based on the value of cChannel
        Select Case cChannel
            Case 0
                grayVal = r
            Case 1
                grayVal = g
            Case 2
                grayVal = b
        End Select
        
        'Assign that gray value to each color channel
        imageData(quickVal, y) = grayVal
        imageData(quickVal + 1, y) = grayVal
        imageData(quickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then MenuGrayscaleSingleChannel = 0 Else MenuGrayscaleSingleChannel = 1
        
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
    ApplyThemeAndTranslations Me
    
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
