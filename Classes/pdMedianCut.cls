VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdMedianCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Median-Cut Color Quantization Class
'Copyright 2017-2018 by Tanner Helland
'Created: 12/January/17
'Last updated: 11/September/17
'Last update: new support for variance and variance + median modes, which are slightly slower but offer better quality
'
'This class provides a highly optimized Median Cut Quantization implementation.  For a nice overview of
' how the median cut algorithm works, see http://www.leptonica.com/papers/mediancut.pdf
'
'For best results, I *strongly* suggest that you resample images to a smaller size before calculating
' an optimized palette via this class.  Specifically, the DIBs.ResizeDIBByPixelCount() function
' allows you to resize an image to a specific number of pixels, which will greatly improve this class's
' performance without harming the final palette's quality.  (In fact, the natural "softening" that comes
' with resampling may actually improve the final palette.)
'
'To use this class, first add all colors from the image via the AddColor_RGB() function.  This class will
' automatically generate related statistical data as colors are added.
'
'After adding all colors, use Split() function to automatically split this class into two new classes.
' Each class will contain roughly half the colors in the image, split along the axis with the largest variance.
'
'Then recursively query the variance of all generated classes (using the .GetVariance() function),
' while always asking the highest-variance class to split itself.  Stop splitting classes when the desired
' number of colors is reached.  (Each class instance maps to one color in the final image.)
'
'To generate the final palette, query each class instance using the GetAverageColor() function.
'
'For a full example of how to use this class, please see the Palettes.GetOptimizedPalette() function.
' That function also demonstrates some caveats, like workarounds for the (impractical) case where a caller
' requests a single-color palette.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Const INITIAL_STACK_SIZE As Long = 256

'Used for testing color weighting by human eye sensitivity.  I haven't made up my mind on the best way to
' handle this.  Using standard illuminants is great for photographs of humans, as it preferentially preserves
' warm tones over cool ones - but blue is weighted *so* lightly that you lose all nuance in nature photography
' involving water or skies.  For now, I'm using reduced versions of the illuminant modifiers, which provides
' a "halfway" solution.
Private Const CUSTOM_WEIGHT_RED As Single = 0.33     '0.299
Private Const CUSTOM_WEIGHT_GREEN As Single = 0.42   '0.587
Private Const CUSTOM_WEIGHT_BLUE As Single = 0.25    '0.114

Private Type RGBStackItem
    Red As Byte
    Green As Byte
    Blue As Byte
    Flag As Byte
    Count As Long
End Type

Public Enum PD_QuantizeMode
    pdqs_MinMax = 0
    pdqs_Variance = 1
    pdqs_VarPlusMedian = 2
End Enum

#If False Then
    Private Const pdqs_MinMax = 0, pdqs_Variance = 1, pdqs_VarPlusMedian = 2
#End If

'As a performance optimization, quantize mode must be specified prior to actually adding pixels to the stack.
' This allows us to calculate statistical data "as we go".
Private m_QuantizeMode As PD_QuantizeMode

Private m_Colors() As RGBStackItem
Private m_NumOfColors As Long       'The number of colors currently stored in the m_Colors() array
Private m_NumOfColorsAdded As Long  'The number of colors the caller has tried to add (always >= m_NumOfColors)
Private m_LastColor As Long

'Some quantization modes can perform calculations "as we go", which tends to be more efficient than calculating
' large batches of data after-the-fact.
Private m_RMin As Long, m_RMax As Long, m_GMin As Long, m_GMax As Long, m_BMin As Long, m_BMax As Long

'Variance-based quantization uses variance (instead of volume) as the criteria for splitting cubes, and it tends to
' provide better results.  To improve performance, we use a neat optimization from a 1962 paper by BP Welford
' (see http://www.johndcook.com/standard_deviation.html); this requires a few extra variables.
Private m_pxVarCount As Long
Private m_MeanR As Double, m_MeanG As Double, m_MeanB As Double
Private m_MedianR As Long, m_MedianG As Long, m_MedianB As Long
Private m_OldMeanR As Double, m_OldMeanG As Double, m_OldMeanB As Double
Private m_VarianceR As Double, m_VarianceG As Double, m_VarianceB As Double
Private m_VarianceUpToDate As Boolean

Friend Function GetQuantizeMode() As PD_QuantizeMode
    GetQuantizeMode = m_QuantizeMode
End Function

Friend Sub SetQuantizeMode(ByVal newMode As PD_QuantizeMode)
    m_QuantizeMode = newMode
End Sub

Friend Sub AddColor_RGB(ByVal srcR As Byte, ByVal srcG As Byte, ByVal srcB As Byte, Optional ByVal cCount As Long = 1)
    
    'Cheap RLE optimization
    If (RGB(srcR, srcG, srcB) <> m_LastColor) Then
        
        m_LastColor = RGB(srcR, srcG, srcB)
        
        With m_Colors(m_NumOfColors)
            .Red = srcR
            .Green = srcG
            .Blue = srcB
            
            'Reset the flag (flags are used with various reduction techniques, so its meaning is conditional)
            .Flag = 0
            .Count = cCount
        End With
        
        m_NumOfColors = m_NumOfColors + 1
        m_NumOfColorsAdded = m_NumOfColorsAdded + cCount
        If (m_NumOfColors > UBound(m_Colors)) Then ReDim Preserve m_Colors(0 To m_NumOfColors * 2 - 1) As RGBStackItem
        
        'Recalculate any running statistical data
        If (m_QuantizeMode = pdqs_MinMax) Then
            
            If (srcR > m_RMax) Then
                m_RMax = srcR
            ElseIf (srcR < m_RMin) Then
                m_RMin = srcR
            End If
            
            If (srcG > m_GMax) Then
                m_GMax = srcG
            ElseIf (srcG < m_GMin) Then
                m_GMin = srcG
            End If
            
            If (srcB > m_BMax) Then
                m_BMax = srcB
            ElseIf (srcB < m_BMin) Then
                m_BMin = srcB
            End If
            
        End If
        
    Else
        m_Colors(m_NumOfColors - 1).Count = m_Colors(m_NumOfColors - 1).Count + cCount
        m_NumOfColorsAdded = m_NumOfColorsAdded + cCount
    End If
    
    'Variance-based quantization requires us to re-calculate variance after modifying the color tables.
    ' Rather than waste a branch on this, just always reset the variance tracker.
    m_VarianceUpToDate = False
    
End Sub

Private Sub CalculateMinMax()

    m_RMin = 256
    m_RMax = -1
    m_GMin = 256
    m_GMax = -1
    m_BMin = 256
    m_BMax = -1
    
    Dim i As Long
    For i = 0 To m_NumOfColors - 1
        With m_Colors(i)
            If (.Red < m_RMin) Then
                m_RMin = .Red
            ElseIf (.Red > m_RMax) Then
                m_RMax = .Red
            End If
            If (.Green < m_GMin) Then
                m_GMin = .Green
            ElseIf (.Green > m_GMax) Then
                m_GMax = .Green
            End If
            If (.Blue < m_BMin) Then
                m_BMin = .Blue
            ElseIf (.Blue > m_BMax) Then
                m_BMax = .Blue
            End If
        End With
    Next i
    
End Sub

Friend Sub GetVariance(ByRef dstRV As Single, ByRef dstGV As Single, ByRef dstBV As Single)
    
    If (m_QuantizeMode = pdqs_MinMax) Then
        dstRV = (m_RMax - m_RMin) * CUSTOM_WEIGHT_RED
        dstGV = (m_GMax - m_GMin) * CUSTOM_WEIGHT_GREEN
        dstBV = (m_BMax - m_BMin) * CUSTOM_WEIGHT_BLUE
    
    Else
    
        If (Not m_VarianceUpToDate) Then UpdateVarianceTrackers
        
        'This stack contains a single color entry, which is okay; that just means it is unsplittable
        If (m_NumOfColors <= 1) Then
            dstRV = 0!
            dstGV = 0!
            dstBV = 0!
        Else
            Dim tmpDivisor As Double
            tmpDivisor = 1# / (CDbl(m_pxVarCount - 1))
            dstRV = m_VarianceR * tmpDivisor
            dstGV = m_VarianceG * tmpDivisor
            dstBV = m_VarianceB * tmpDivisor
        End If
        
    End If
    
End Sub

Friend Function GetNumOfColors() As Long
    GetNumOfColors = m_NumOfColors
End Function

Friend Sub Reset()

    ReDim m_Colors(0 To INITIAL_STACK_SIZE - 1) As RGBStackItem
    m_NumOfColors = 0
    m_NumOfColorsAdded = 0
    m_LastColor = -1
    
    m_RMin = 256
    m_RMax = -1
    m_GMin = 256
    m_GMax = -1
    m_BMin = 256
    m_BMax = -1
    
    m_VarianceUpToDate = False
    
End Sub

'Split (roughly) half of this stack into some new stack, using the criteria specified.
Friend Sub Split(ByRef dstStack As pdMedianCut)
    
    If (dstStack Is Nothing) Then Set dstStack = New pdMedianCut
    dstStack.SetQuantizeMode Me.GetQuantizeMode()
    
    'Find the channel with maximum variance
    Dim rDiff As Single, gDiff As Single, bDiff As Single
    Me.GetVariance rDiff, gDiff, bDiff
        
    If (m_QuantizeMode = pdqs_MinMax) Then
        
        'Apply the split.  Note that all split functions set the flag of removed items to (1) to note that
        ' the pixel is no longer part of this stack.
        If (rDiff > gDiff) Then
            If (rDiff > bDiff) Then SplitR dstStack, (m_RMax + m_RMin) \ 2 Else SplitB dstStack, (m_BMax + m_BMin) \ 2
        Else
            If (gDiff > bDiff) Then SplitG dstStack, (m_GMax + m_GMin) \ 2 Else SplitB dstStack, (m_BMax + m_BMin) \ 2
        End If
    
    ElseIf (m_QuantizeMode = pdqs_Variance) Then
        
        'Split by mean:
        If (rDiff > gDiff) Then
            If (rDiff > bDiff) Then SplitR dstStack, Int(m_MeanR) Else SplitB dstStack, Int(m_MeanB)
        Else
            If (gDiff > bDiff) Then SplitG dstStack, Int(m_MeanG) Else SplitB dstStack, Int(m_MeanB)
        End If

    ElseIf (m_QuantizeMode = pdqs_VarPlusMedian) Then
    
        'Split by median:
        If (rDiff > gDiff) Then
            If (rDiff > bDiff) Then SplitR dstStack, m_MedianR Else SplitB dstStack, m_MedianB
        Else
            If (gDiff > bDiff) Then SplitG dstStack, m_MedianG Else SplitB dstStack, m_MedianB
        End If
        
    End If
    
    'After a split, purge any/all removed items from the stack
    RemoveFlaggedEntries
    
End Sub

'Generic split options, separated by channel.  The caller is responsible for determining a midpoint.
Private Sub SplitR(ByRef dstStack As pdMedianCut, ByVal rMidpoint As Long)
    Dim i As Long
    For i = 0 To m_NumOfColors - 1
        If (m_Colors(i).Red > rMidpoint) Then
            With m_Colors(i)
                dstStack.AddColor_RGB .Red, .Green, .Blue, .Count
                .Flag = 1
            End With
        End If
    Next i
End Sub

Private Sub SplitG(ByRef dstStack As pdMedianCut, ByVal gMidpoint As Long)
    Dim i As Long
    For i = 0 To m_NumOfColors - 1
        If (m_Colors(i).Green > gMidpoint) Then
            With m_Colors(i)
                dstStack.AddColor_RGB .Red, .Green, .Blue, .Count
                .Flag = 1
            End With
        End If
    Next i
End Sub

Private Sub SplitB(ByRef dstStack As pdMedianCut, ByVal bMidpoint As Long)
    Dim i As Long
    For i = 0 To m_NumOfColors - 1
        If (m_Colors(i).Blue > bMidpoint) Then
            With m_Colors(i)
                dstStack.AddColor_RGB .Red, .Green, .Blue, .Count
                .Flag = 1
            End With
        End If
    Next i
End Sub

'Return the average color of this stack.  Note that "average" always means "weighted average," as colors are weighted
' by their prominence in the original image.
Friend Sub GetAverageColor(ByRef r As Long, ByRef g As Long, ByRef b As Long)
    
    If (m_NumOfColors > 0) Then
        
        If (m_QuantizeMode = pdqs_MinMax) Then
            
            Dim rTotal As Long, gTotal As Long, bTotal As Long
            Dim i As Long
            For i = 0 To m_NumOfColors - 1
                With m_Colors(i)
                    rTotal = rTotal + .Red * .Count
                    gTotal = gTotal + .Green * .Count
                    bTotal = bTotal + .Blue * .Count
                End With
            Next i
        
            r = rTotal \ m_NumOfColorsAdded
            g = gTotal \ m_NumOfColorsAdded
            b = bTotal \ m_NumOfColorsAdded
            
        ElseIf (m_QuantizeMode = pdqs_Variance) Then
        
            If (Not m_VarianceUpToDate) Then UpdateVarianceTrackers
            
            r = Int(m_MeanR + 0.5)
            g = Int(m_MeanG + 0.5)
            b = Int(m_MeanB + 0.5)
            
        ElseIf (m_QuantizeMode = pdqs_VarPlusMedian) Then
            
            If (Not m_VarianceUpToDate) Then UpdateVarianceTrackers
            
            'Even though we split according to the median value, we still report average color
            ' using mean values.  (Specifically, *weighted* mean values.)
            r = Int(m_MeanR + 0.5)
            g = Int(m_MeanG + 0.5)
            b = Int(m_MeanB + 0.5)
            
        End If
        
    End If
    
End Sub

Private Sub UpdateVarianceTrackers()
    
    m_VarianceR = 0#
    m_VarianceG = 0#
    m_VarianceB = 0#
    
    'Histogram code is optional, but it yields better split points
    If (m_QuantizeMode = pdqs_VarPlusMedian) Then
        Dim rHist() As Long, gHist() As Long, bHist() As Long
        ReDim rHist(0 To 255) As Long: ReDim gHist(0 To 255) As Long: ReDim bHist(0 To 255) As Long
    End If
    
    If (m_NumOfColors > 0) Then
    
        'Find the first unflagged entry.  (NOTE: this is no longer necessary, as flagged entries are automatically removed.)
        Dim i As Long
        
        m_OldMeanR = m_Colors(0).Red
        m_OldMeanG = m_Colors(0).Green
        m_OldMeanB = m_Colors(0).Blue
        
        m_MeanR = m_OldMeanR
        m_MeanG = m_OldMeanG
        m_MeanB = m_OldMeanB
        
        m_pxVarCount = m_Colors(0).Count
        
        If (m_QuantizeMode = pdqs_VarPlusMedian) Then
            rHist(m_Colors(0).Red) = m_pxVarCount
            gHist(m_Colors(0).Green) = m_pxVarCount
            bHist(m_Colors(0).Blue) = m_pxVarCount
        End If
        
        'Calculate all other colors in turn
        If (m_NumOfColors > 1) Then
            
            Dim tmpRed As Double, tmpGreen As Double, tmpBlue As Double
            Dim dDeltaR As Double, dDeltaG As Double, dDeltaB As Double
            Dim myCount As Double, combinedCount As Double, dPxCount As Double
            
            For i = 1 To m_NumOfColors - 1
                
                With m_Colors(i)
                
                    tmpRed = CDbl(.Red)
                    tmpGreen = CDbl(.Green)
                    tmpBlue = CDbl(.Blue)
                    
                    'Calculate deltas
                    dDeltaR = tmpRed - m_MeanR
                    dDeltaG = tmpGreen - m_MeanG
                    dDeltaB = tmpBlue - m_MeanB
                    dDeltaR = dDeltaR * dDeltaR
                    dDeltaG = dDeltaG * dDeltaG
                    dDeltaB = dDeltaB * dDeltaB
                    
                    'Combined count is only ever used as a divisor
                    myCount = .Count
                    dPxCount = m_pxVarCount
                    combinedCount = 1# / (dPxCount + myCount)
                    
                    'Calculate new means (by basically weighting this entry against its pixel count)
                    m_MeanR = ((m_OldMeanR * dPxCount) + (tmpRed * myCount)) * combinedCount
                    m_MeanG = ((m_OldMeanG * dPxCount) + (tmpGreen * myCount)) * combinedCount
                    m_MeanB = ((m_OldMeanB * dPxCount) + (tmpBlue * myCount)) * combinedCount
                    
                    'Calculate new variances (and again, weight them by the pixel count of this color)
                    m_VarianceR = m_VarianceR + (dDeltaR * dPxCount * myCount) * combinedCount
                    m_VarianceG = m_VarianceG + (dDeltaG * dPxCount * myCount) * combinedCount
                    m_VarianceB = m_VarianceB + (dDeltaB * dPxCount * myCount) * combinedCount
                    
                    'Update the running pixel count and averages
                    m_OldMeanR = m_MeanR
                    m_OldMeanG = m_MeanG
                    m_OldMeanB = m_MeanB
                    m_pxVarCount = m_pxVarCount + myCount
                    
                    'Also update the histograms as necessary
                    If (m_QuantizeMode = pdqs_VarPlusMedian) Then
                        rHist(.Red) = rHist(.Red) + myCount
                        gHist(.Green) = gHist(.Green) + myCount
                        bHist(.Blue) = bHist(.Blue) + myCount
                    End If
                    
                End With
                
            Next i
            
        Else
            'This stack only contains one color, which means it cannot be split.  That's okay!
            'Debug.Print "NOTE!  m_NumOfColors = 1"
        End If
    
    Else
        'This stack is empty.  That typically only happens when trying to reduce an image to n colors,
        ' but the image contains less than n colors - meaning stacks cannot be split.  That's okay!
        'Debug.Print "WARNING!  m_NumOfColors = 0"
    End If
    
    'Variance can end up being quite small (due to rounding errors), so if it falls below a certain threshold,
    ' treat this stack as if it only contains *1* color.
    If (m_VarianceR < 0.00001) Then m_VarianceR = 0#
    If (m_VarianceG < 0.00001) Then m_VarianceG = 0#
    If (m_VarianceB < 0.00001) Then m_VarianceB = 0#
    If (m_VarianceR = 0) And (m_VarianceG = 0) And (m_VarianceB = 0) Then m_NumOfColors = 1
    
    'Calculate medians as necessary
    If (m_QuantizeMode = pdqs_VarPlusMedian) Then
        
        'Ignore stacks with only one color
        If (m_NumOfColors > 1) Then
        
            Dim histSort(0 To 255) As Long
            Dim curCount As Long, curIndex As Long
            
            curCount = 0: curIndex = -1
            Do
                curIndex = curIndex + 1
                If (rHist(curIndex) <> 0) Then
                    histSort(curCount) = curIndex
                    curCount = curCount + 1
                End If
            Loop While (curIndex < 255)
            m_MedianR = histSort((curCount - 1) \ 2)
            
            curCount = 0: curIndex = -1
            Do
                curIndex = curIndex + 1
                If (gHist(curIndex) <> 0) Then
                    histSort(curCount) = curIndex
                    curCount = curCount + 1
                End If
            Loop While (curIndex < 255)
            m_MedianG = histSort((curCount - 1) \ 2)
            
            curCount = 0: curIndex = -1
            Do
                curIndex = curIndex + 1
                If (bHist(curIndex) <> 0) Then
                    histSort(curCount) = curIndex
                    curCount = curCount + 1
                End If
            Loop While (curIndex < 255)
            m_MedianB = histSort((curCount - 1) \ 2)
            
        Else
            m_MedianR = m_Colors(0).Red
            m_MedianG = m_Colors(0).Green
            m_MedianB = m_Colors(0).Blue
        End If
        
    End If
    
    'Variance is now up-to-date for this cube
    m_VarianceUpToDate = True

End Sub

Friend Sub CopyStackToRGBQuad(ByRef dstArray() As RGBQuad)
    ReDim dstArray(0 To m_NumOfColors - 1) As RGBQuad
    Dim i As Long
    For i = 0 To m_NumOfColors - 1
        dstArray(i).Red = m_Colors(i).Red
        dstArray(i).Green = m_Colors(i).Green
        dstArray(i).Blue = m_Colors(i).Blue
    Next i
End Sub

Private Sub RemoveFlaggedEntries()

    Dim newOffset As Long
    Dim i As Long
    For i = 0 To m_NumOfColors - 1
        If (m_Colors(i).Flag <> 0) Then
            newOffset = newOffset + 1
            m_NumOfColorsAdded = m_NumOfColorsAdded - m_Colors(i).Count
        Else
            If (newOffset <> 0) Then m_Colors(i - newOffset) = m_Colors(i)
        End If
    Next i
    
    m_NumOfColors = m_NumOfColors - newOffset
    
    'Certain quantize modes may need to perform certain calculation maintenance tasks after a purge
    If (m_QuantizeMode = pdqs_MinMax) Then CalculateMinMax
    m_VarianceUpToDate = False
    
End Sub

Private Sub Class_Initialize()
    Me.Reset
    m_QuantizeMode = pdqs_MinMax
    m_VarianceUpToDate = False
End Sub
