VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdSurfaceF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon 32-bit-float-per-channel surface
'Copyright 2023-2026 by Tanner Helland
'Created: 08/January/23
'Last updated: 10/January/23
'Last update: continue building out feature set
'
'This class manages pixel data for a surface in 32-bit, single-precision float format.
' By default, colors are assumed to be on the range [0, 1] but HDR values are of course supported too!
' Just be aware of the nuances involved when converting such data back to 32-bit (or other) formats.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Surface width and height.  0 for either indicates an uninitialized surface.
Private m_Width As Long, m_Height As Long

'Number of channels.  Currently supports 3- and 4-channel color data; grayscale would be nice to support (someday).
Private m_NumChannels As Long

'Stride is calculated as width * channels, and is filled automatically at initialization time
Private m_Stride As Long

'Alpha premultiplication state
Private m_IsAlphaPremultiplied As Boolean

'Actual pixel data, as single-precision floats.
Private m_Pixels() As Single

'I will likely rework this in the future, but for now, let's declare a single class-level type for temporary usage.
Private m_tmpData() As Byte

'Ensure color and alpha values are locked to the [0, 1] range.
Friend Sub ClampGamut()
    
    'Basic validation
    If (m_Width = 0) Or (m_Height = 0) Or (m_NumChannels = 0) Then Exit Sub
    
    Dim x As Long, y As Long, xMax As Long, yMax As Long
    xMax = m_Stride - 1 'Note the use of stride, not width
    yMax = m_Height - 1
    
    For y = 0 To yMax
    For x = 0 To xMax
        If (m_Pixels(x, y) < 0!) Then m_Pixels(x, y) = 0!
        If (m_Pixels(x, y) > 1!) Then m_Pixels(x, y) = 1!
    Next x
    Next y
    
End Sub

'Quickly convert this surface to a 24/32-bit destination pdDIB.  Linear gamma (1.0) is assumed; I'll likely expand
' gamma handling in the future, as relevant.
Friend Function ConvertToPDDib(ByRef dstDIB As pdDIB) As Boolean
    
    ConvertToPDDib = False
    On Error GoTo CouldNotConvertSurface
    
    'Basic validation
    If (m_Width = 0) Or (m_Height = 0) Or (m_NumChannels = 0) Then Exit Function
    
    'Initialize a new DIB of matching size and depth
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    ConvertToPDDib = dstDIB.CreateBlank(m_Width, m_Height, m_NumChannels * 8, 0, 0)
    If (Not ConvertToPDDib) Then Exit Function
    
    Dim x As Long, y As Long, z As Long, xStride As Long, maxChannel As Long, tmpColor As Single, tmpByte As Long
    maxChannel = m_NumChannels - 1
    
    'Get direct access to the source DIB's pixel data
    Dim dstPixels() As Byte, dstPixelSA As SafeArray1D
    Dim xFinal As Long, yFinal As Long
    xFinal = m_Width - 1
    yFinal = m_Height - 1
    
    'Manually copy bytes over.  (TODO: look at a 3rd-party library like lcms2 for better perf?)
    For y = 0 To yFinal
        
        dstDIB.WrapArrayAroundScanline dstPixels, dstPixelSA, y
        
        'Generate target RGB/A values
        For x = 0 To xFinal
            
            xStride = x * m_NumChannels
            
            For z = 0 To maxChannel
                
                tmpColor = m_Pixels(xStride + z, y)
                If (tmpColor < 0!) Then tmpColor = 0!
                If (tmpColor > 1!) Then tmpColor = 1!
                
                'Improve distribution across all [0, 255] values by using 255.9999
                ' (while also correctly round-tripping original [0, 255] byte values).
                dstPixels(xStride + z) = Int(tmpColor * 255.9999!)
                
            Next z
            
            'Swizzle R/B to match Windows BGRA order
            tmpByte = dstPixels(xStride)
            dstPixels(xStride) = dstPixels(xStride + 2)
            dstPixels(xStride + 2) = tmpByte
            
        Next x
        
    Next y
    
    'Free unsafe array wrappers
    dstDIB.UnwrapArrayFromDIB dstPixels
    
    'Set any other state parameters
    dstDIB.SetInitialAlphaPremultiplicationState m_IsAlphaPremultiplied
    
    ConvertToPDDib = True
    Exit Function

'The only expected error in this function would be an out-of-memory error; nothing we can do if that occurs
CouldNotConvertSurface:
    
    Me.EraseSurface
    ConvertToPDDib = False
    
End Function

'Create a blank float surface.  Channel count is optional, with a default of 3-channel RGB.
' You can also set per-channel initialization values.  By default, all channels are initialized to 0 (transparent black).
Friend Function CreateBlank(ByVal imgWidth As Long, ByVal imgHeight As Long, Optional ByVal channelCount As Long = 3, Optional ByVal initR As Single = 0!, Optional ByVal initG As Single = 0!, Optional ByVal initB As Single = 0!, Optional ByVal initA As Single = 0!) As Boolean
    
    CreateBlank = False
    On Error GoTo CouldNotCreateSurface
    
    'Failsafe checks
    If (imgWidth <= 0) Or (imgHeight <= 0) Or (channelCount < 3) Or (channelCount > 4) Then Exit Function
    
    'Store all passed values permanently and calculate related values
    m_Width = imgWidth
    m_Height = imgHeight
    m_NumChannels = channelCount
    m_Stride = m_Width * channelCount
    
    'In the future, we will optimize this to detect current memory allocations and reuse them if possible.
    ' For now, however, creating a blank surface will always re-create all structures from scratch.
    ReDim m_Pixels(0 To m_Stride - 1, 0 To m_Height - 1) As Single
    
    'Next, apply colors.  To improve performance, we'll prepare a single color entry and simply copy that
    ' color across all pixels in the image.
    If (initR <> 0!) Or (initG <> 0!) Or (initB <> 0!) Or (initA <> 0!) Then
    
        'Allocate a temporary color "holder"
        Dim pxSizeBytes As Long
        pxSizeBytes = m_NumChannels * 4
        ReDim m_tmpData(0 To pxSizeBytes - 1) As Byte
        
        'Populate with passed colors
        GetMem4_Ptr VarPtr(initR), VarPtr(m_tmpData(0))
        GetMem4_Ptr VarPtr(initG), VarPtr(m_tmpData(4))
        GetMem4_Ptr VarPtr(initB), VarPtr(m_tmpData(8))
        If (m_NumChannels = 4) Then GetMem4_Ptr VarPtr(initA), VarPtr(m_tmpData(12))
        
        Dim pTmpColor As Long
        pTmpColor = VarPtr(m_tmpData(0))
        
        'Iterate through the new image and fill it with the placeholder color
        Dim x As Long, y As Long
        For y = 0 To imgHeight - 1
        For x = 0 To imgWidth - 1
            CopyMemoryStrict VarPtr(m_Pixels(x * pxSizeBytes, y)), pTmpColor, pxSizeBytes
        Next x
        Next y
        
    End If
    
    'Auto-detect common alpha premultiplication combinations
    m_IsAlphaPremultiplied = False
    If (initR = 0!) And (initG = 0!) And (initB = 0!) And (initA = 0!) Then
        m_IsAlphaPremultiplied = True
    ElseIf (initR = 1!) And (initG = 1!) And (initB = 1!) And (initA = 1!) Then
        m_IsAlphaPremultiplied = True
    End If
    
    CreateBlank = True
    Exit Function

'The only expected error in this function would be an out-of-memory error; nothing we can do if that occurs
CouldNotCreateSurface:
    
    Me.EraseSurface
    CreateBlank = False
    
End Function

'Quick-create a new 96/128-bit float surface from a 24/32-bit pdDIB
Friend Function CreateFromPDDib(ByRef srcDIB As pdDIB) As Boolean
    
    CreateFromPDDib = False
    On Error GoTo CouldNotCreateSurface
    
    'Basic validation
    If (srcDIB Is Nothing) Then Exit Function
    If (srcDIB.GetDIBWidth = 0) Or (srcDIB.GetDIBHeight = 0) Then Exit Function
    If (srcDIB.GetDIBPointer = 0) Then Exit Function
    
    'Initialize a new float surface of matching size and depth
    CreateFromPDDib = Me.CreateBlank(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBColorDepth \ 8)
    If (Not CreateFromPDDib) Then
        Me.EraseSurface
        Exit Function
    End If
    
    Dim x As Long, y As Long, xStride As Long
    
    'Conversion from [0, 255] to [0.0, 1.0] is faster with a fixed lookup table
    Const BYTE_TO_FLOAT As Single = 1! / 255!
    Dim pxByteToFloat(0 To 255) As Single
    For x = 0 To 255
        pxByteToFloat(x) = x * BYTE_TO_FLOAT
    Next x
    
    'Get direct access to the source DIB's pixel data
    Dim srcPixels() As RGBQuad, srcPixelSA As SafeArray1D
    
    'Manually copy bytes over.  (TODO: look at a 3rd-party library like lcms2 for better perf?)
    For y = 0 To m_Height - 1
        srcDIB.WrapRGBQuadArrayAroundScanline srcPixels, srcPixelSA, y
    For x = 0 To m_Width - 1
        
        xStride = x * m_NumChannels
        
        With srcPixels(x)
            m_Pixels(xStride, y) = pxByteToFloat(.Red)
            m_Pixels(xStride + 1, y) = pxByteToFloat(.Green)
            m_Pixels(xStride + 2, y) = pxByteToFloat(.Blue)
            If (m_NumChannels = 4) Then m_Pixels(xStride + 3, y) = pxByteToFloat(.Alpha)
        End With
        
    Next x
    Next y
    
    'Free unsafe array wrappers
    srcDIB.UnwrapRGBQuadArrayFromDIB srcPixels
    
    'Mirror any other passed state values
    m_IsAlphaPremultiplied = srcDIB.GetAlphaPremultiplication()
    
    CreateFromPDDib = True
    Exit Function

'The only expected error in this function would be an out-of-memory error; nothing we can do if that occurs
CouldNotCreateSurface:
    
    Me.EraseSurface
    CreateFromPDDib = False
    
End Function

'Create a new surface from data at an arbitrary pointer.  Pointer safety should be handled by the caller,
' but an optional data length can be passed to have this function handle safety.
Friend Function CreateFromPtr(ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal channelCount As Long, ByVal srcDataPtr As Long, Optional ByVal srcDataLen As Long = 0, Optional ByVal isAlphaPremultiplied As Boolean = False) As Boolean
    
    CreateFromPtr = False
    
    'Validate passed pointer data
    If (srcDataPtr = 0) Then Exit Function
    If (srcDataLen <> 0) Then
        If (srcDataLen <> imgWidth * channelCount * 4 * imgHeight) Then Exit Function
    Else
        srcDataLen = imgWidth * channelCount * 4 * imgHeight
    End If
    
    'Create a blank, black transparent image (this skips any manual color initialization).
    ' Note that this step will also validate passed width/height/channels.
    CreateFromPtr = Me.CreateBlank(imgWidth, imgHeight, channelCount, 0!, 0!, 0!, 0!)
    
    'Copy the source pixels into the new buffer
    If CreateFromPtr Then CopyMemoryStrict VarPtr(m_Pixels(0, 0)), srcDataPtr, srcDataLen
    
    'Mirror any other passed state values
    m_IsAlphaPremultiplied = isAlphaPremultiplied
    
End Function

Friend Sub EraseSurface()
    Erase m_Pixels
    Erase m_tmpData
    m_Width = 0
    m_Height = 0
    m_Stride = 0
    m_NumChannels = 0
End Sub

'Get/set alpha premultiplication.
' IMPORTANT NOTE!  To make it explicitly clear that modifying this property DOES NOT ACTUALLY MODIFY PIXEL DATA,
' the Set instruction is labeled differently.  It is only meant to be used by surface creation functions,
' where the premultiplication state is explicitly known prior to writing pixel data.  The counterpart
' SetAlphaPremultiplication function will actually *modify* pixels as necessary to create the desired
' premultiplication state.
Friend Function GetAlphaPremultiplication() As Boolean
    GetAlphaPremultiplication = m_IsAlphaPremultiplied
End Function

Friend Sub SetInitialAlphaPremultiplicationState(ByVal newState As Boolean)
    m_IsAlphaPremultiplied = newState
End Sub

Friend Function GetChannelCount() As Long
    GetChannelCount = m_NumChannels
End Function

Friend Function GetHeight() As Long
    GetHeight = m_Height
End Function

Friend Function GetPixelPtr() As Long
    If (m_Width <> 0) Then GetPixelPtr = VarPtr(m_Pixels(0, 0))
End Function

Friend Function GetPixelSize() As Long
    GetPixelSize = m_Width * m_NumChannels * 4 * m_Height
End Function

Friend Function GetStride() As Long
    GetStride = m_Stride
End Function

Friend Function GetWidth() As Long
    GetWidth = m_Width
End Function

Friend Sub Reset()
    If (Me.GetPixelPtr <> 0) Then VBHacks.ZeroMemory Me.GetPixelPtr, m_Stride * m_Height * 4
End Sub

'This routine can be used to either apply or remove premultiplied alpha values.
Friend Function SetAlphaPremultiplication(Optional ByVal applyPremultiplication As Boolean = False, Optional ByVal ignoreEmbeddedValue As Boolean = False) As Boolean
    
    SetAlphaPremultiplication = False
    
    'Validate basic settings
    If (Me.GetPixelSize <= 0) Then Exit Function
    If (m_NumChannels <> 4) Then Exit Function
    
    'If alpha premultiplication already matches the requested state, exit now
    If (m_IsAlphaPremultiplied = applyPremultiplication) And (Not ignoreEmbeddedValue) Then Exit Function
    m_IsAlphaPremultiplied = applyPremultiplication
    
    Dim x As Long, y As Long, xStride As Long, curAlpha As Single
    
    Dim xFinal As Long, yFinal As Long
    xFinal = m_Width - 1
    yFinal = m_Height - 1
    
    'Multiply colors by their alpha component
    If applyPremultiplication Then
        
        For y = 0 To yFinal
        For x = 0 To xFinal
            
            xStride = x * 4
            curAlpha = m_Pixels(xStride + 3, y)
            
            m_Pixels(xStride, y) = m_Pixels(xStride, y) * curAlpha
            m_Pixels(xStride + 1, y) = m_Pixels(xStride + 1, y) * curAlpha
            m_Pixels(xStride + 2, y) = m_Pixels(xStride + 2, y) * curAlpha
            
        Next x
        Next y
        
        SetAlphaPremultiplication = True
    
    'Divide colors by their alpha component
    Else
        
        For y = 0 To yFinal
        For x = 0 To xFinal
            
            xStride = x * 4
            curAlpha = m_Pixels(xStride + 3, y)
            
            If (curAlpha <> 0!) Then
                curAlpha = 1! / curAlpha
                m_Pixels(xStride, y) = m_Pixels(xStride, y) * curAlpha
                m_Pixels(xStride + 1, y) = m_Pixels(xStride + 1, y) * curAlpha
                m_Pixels(xStride + 2, y) = m_Pixels(xStride + 2, y) * curAlpha
            End If
            
        Next x
        Next y
        
        SetAlphaPremultiplication = True
        
    End If
    
End Function

'Convenience functions for wrapping an array around this DIB's bits.
' You *must* call the Unwrap function prior to the array falling out of scope, or VB will crash.
Friend Sub WrapArrayAroundSurface(ByRef srcArray() As Single, ByRef srcSafeArray As SafeArray2D)
    PrepInternalSafeArray srcSafeArray
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub WrapArrayAroundSurface_1D(ByRef srcArray() As Single, ByRef srcSafeArray As SafeArray1D)
    PrepInternalSafeArray_Scanline srcSafeArray, 0
    srcSafeArray.cElements = m_Stride * m_Height
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub WrapArrayAroundScanline(ByRef srcArray() As Single, ByRef srcSafeArray As SafeArray1D, Optional ByVal dstScanLine As Long = 0)
    PrepInternalSafeArray_Scanline srcSafeArray, dstScanLine
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub UnwrapArrayFromSurface(ByRef srcArray() As Single)
    PutMem4 VarPtrArray(srcArray), 0&
End Sub

'These functions allow for unsafe direct access to the underlying pixel data.
' Note that this function *will* produce crashes if the underlying surface is *not* initialized,
' but we do not check that case here for performance reasons.  (The caller must validate this manually.)
Private Sub PrepInternalSafeArray(ByRef dstSafeArray As SafeArray2D)
    
    With dstSafeArray
        .cbElements = 4
        .cDims = 2
        .cLocks = 1
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = m_Stride
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = m_Height
        .pvData = VarPtr(m_Pixels(0, 0))
    End With
    
End Sub

Private Sub PrepInternalSafeArray_Scanline(ByRef dstSafeArray As SafeArray1D, ByVal dstLine As Long)
    
    With dstSafeArray
        .cbElements = 4
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = m_Stride
        .pvData = VarPtr(m_Pixels(0, 0)) + (dstLine * m_Stride * 4)
    End With
    
End Sub
