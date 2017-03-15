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
'Copyright 2017-2017 by Tanner Helland
'Created: 12/January/17
'Last updated: 15/January/17
'Last update: weight variance calculations by channel, which meaningfully improves the final palette
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
'At present, this class is wired up for different variance-calculation modes.  However, only the PDQS_MINMAX
' mode is implemented.  (The PDQS_WU enum would theoretically use histograms and true running variance to
' calculate split positions.)  I will also look at adding Lab color-space methods in the future.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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

Public Enum PD_QUANTIZE_MODE
    PDQS_MINMAX = 0
    PDQS_WU = 1
End Enum

#If False Then
    Private Const PDQS_MINMAX = 0, PDQS_WU = 1
#End If

'As a performance optimzation, quantize mode must be specified prior to actually adding pixels to the stack.
' This allows us to calculate statistical data "as we go", which is much faster than calculating it on-the-fly.
Private m_QuantizeMode As PD_QUANTIZE_MODE

Private m_Colors() As RGBStackItem
Private m_NumOfColors As Long       'The number of colors currently stored in the m_Colors() array
Private m_NumOfColorsAdded As Long  'The number of colors the caller has tried to add (always >= m_NumOfColors)
Private m_LastColor As Long

'Some quantization modes can perform calculations "as we go", which tends to be more efficient than calculating
' large batches of data after-the-fact.
Private m_RMin As Long, m_RMax As Long, m_GMin As Long, m_GMax As Long, m_BMin As Long, m_BMax As Long

Private Declare Sub CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal dstPointer As Long, ByVal srcPointer As Long, ByVal numOfBytes As Long)
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Long) As Long

Friend Function GetQuantizeMode() As PD_QUANTIZE_MODE
    GetQuantizeMode = m_QuantizeMode
End Function

Friend Sub SetQuantizeMode(ByVal newMode As PD_QUANTIZE_MODE)
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
        If (m_QuantizeMode = PDQS_MINMAX) Then
            
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

Friend Sub GetMinMax(ByRef dstRMin As Long, ByRef dstRMax As Long, ByRef dstGMin As Long, ByRef dstGMax As Long, ByRef dstBMin As Long, ByRef dstBMax As Long)
    dstRMin = m_RMin
    dstRMax = m_RMax
    dstGMin = m_GMin
    dstGMax = m_GMax
    dstBMin = m_BMin
    dstBMax = m_BMax
End Sub

Friend Sub GetVariance(ByRef dstRV As Single, ByRef dstGV As Single, ByRef dstBV As Single)
    If (m_QuantizeMode = PDQS_MINMAX) Then
        dstRV = (m_RMax - m_RMin) * CUSTOM_WEIGHT_RED
        dstGV = (m_GMax - m_GMin) * CUSTOM_WEIGHT_GREEN
        dstBV = (m_BMax - m_BMin) * CUSTOM_WEIGHT_BLUE
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
    
End Sub

'Split (roughly) half of this stack into some new stack, using the criteria specified.
Friend Sub Split(ByRef dstStack As pdMedianCut)
    
    If (dstStack Is Nothing) Then Set dstStack = New pdMedianCut
    
    If (m_QuantizeMode = PDQS_MINMAX) Then
        
        'Find the channel with maximum variance
        Dim rDiff As Single, gDiff As Single, bDiff As Single
        Me.GetVariance rDiff, gDiff, bDiff
        
        'Apply the split.  Note that all split functions set the flag of removed items to (1) to note that
        ' the pixel is no longer part of this stack.
        If (rDiff > gDiff) Then
            If (rDiff > bDiff) Then SplitR dstStack, (m_RMax + m_RMin) \ 2 Else SplitB dstStack, (m_BMax + m_BMin) \ 2
        Else
            If (gDiff > bDiff) Then SplitG dstStack, (m_GMax + m_GMin) \ 2 Else SplitB dstStack, (m_BMax + m_BMin) \ 2
        End If
    
    ElseIf (m_QuantizeMode = PDQS_WU) Then
    
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
        
        If (m_QuantizeMode = PDQS_MINMAX) Then
            
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
            
        End If
        
    End If
    
End Sub

Friend Sub CopyStackToRGBQuad(ByRef dstArray() As RGBQUAD)
    ReDim dstArray(0 To m_NumOfColors - 1) As RGBQUAD
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
    If (m_QuantizeMode = PDQS_MINMAX) Then
        CalculateMinMax
    End If

End Sub

Private Sub Class_Initialize()
    Me.Reset
    m_QuantizeMode = PDQS_MINMAX
End Sub
