VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFloodFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Flood Fill Engine
'Copyright 2014-2026 by Tanner Helland
'Created: 11/October/14
'Last updated: 11/March/22
'Last update: fix half-pixel offset when antialiasing is enabled; see https://github.com/tannerhelland/PhotoDemon/issues/395
'
'This class supplies the flood fill algorithm for a number of different tools in PD (magic wand, bucket fill, etc).
' The flood fill approach itself is pretty basic - a stack is used in place of recursion, and a global mapping array
' is used to minimize the amount of pixel checks that take place.
'
'As a convenience to calling functions, this class exposes a number of options.  Fills can be contiguous (default)
' or global, meaning the entire image is searched without regard to continuity.  The comparison mode used between
' pixels can also be specified; individual channels (including alpha), luminosity, or a full composite check of all
' channels can be specified.  Finally, antialiasing can also be requested.  A custom AA solution is used for
' contiguous fills, and it's extremely fast as we know in advance which pixels should be examined for AA.  For
' global fills, PD's standard QuickBlur function is used (as we don't have a continuity map available).
'
'To allow this class to be used by any external function, it simply requires a source and destination DIB.
' Both source and destination DIBs *must be 32-bpp*.  The results of the fill will be placed inside the 32-bpp image
' as grayscale+alpha data, which makes it very easy to render (or apply) any operation based off the floodfill data.
'
'Similarly, how the caller uses the fill map is up to them.  In the case of magic wand selections, PD converts the
' flood fill map to a selection map.  For bucket fill, it simply merges the requested fill type onto the image,
' using the fill map as a guide.
'
'Like any array-based tool, this class will be slow inside the IDE.  Please use only when compiled.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'A stack is used to track pixels that need to be checked
Private m_Stack() As PointAPI
Private m_StackPosition As Long
Private m_StackHeight As Long
Private Const INITIAL_STACK_HEIGHT As Long = 4096

'To reduce iterations, this function tracks pixels that have already been added to the stack.  The boundary values are
' 0-based, so for a 10x10 image, these values will be (9, 9), respectively.
Private m_BoundsX As Long, m_BoundsY As Long
Private m_AlreadyChecked() As Byte

'The completed results of the fill are stored in a dedicated array; this accelerates certain post-processing tasks.
Private m_FillResults() As Byte

'If outlines are requested, we'll use a pdEdgeDetector instance (plus some helper functions) to generate fill outlines.
Private m_FillOutline As pdEdgeDetector
Private m_OutlineBoundsX As Long, m_OutlineBoundsY As Long
Private m_OutlineCopy() As Byte

'Tolerance allows the user to control the strength of the flood
Private m_Tolerance As Double

'Different compare modes can be used to obtain better results.
Public Enum PD_FloodCompare
    pdfc_Color = 0
    pdfc_Composite = 1
    pdfc_Luminance = 2
    pdfc_Red = 3
    pdfc_Green = 4
    pdfc_Blue = 5
    pdfc_Alpha = 6
End Enum

#If False Then
    Private Const pdfc_Color = 0, pdfc_Composite = 1, pdfc_Luminance = 2, pdfc_Red = 3, pdfc_Green = 4, pdfc_Blue = 5, pdfc_Alpha = 6
#End If

Private m_CompareMode As PD_FloodCompare

'Two search methods are supported: contiguous region based on the initial point (default behavior), or the full image
Public Enum PD_FloodSearch
    pdfs_Contiguous = 0
    pdfs_WholeImage = 1
End Enum

#If False Then
    Private Const pdfs_Contiguous = 0, pdfs_WholeImage = 1
#End If

Private m_SearchMode As PD_FloodSearch

'A custom antialiasing technique can be used to soften the floodfill results
Private m_AntialiasingMode As Boolean

'Obviously the function needs a starting x/y position
Private m_InitPoint As PointAPI

'Flood fills always return a 32-bpp DIB that describes the flood fill results.  This DIB is painted via pd2D.
Private m_Brush As pd2DBrush

'Get/set functions for all relevant flood fill parameters
Friend Function GetAntialiasingMode() As Boolean
    GetAntialiasingMode = m_AntialiasingMode
End Function

Friend Sub SetAntialiasingMode(ByVal newAntialiasingMode As Boolean)
    m_AntialiasingMode = newAntialiasingMode
End Sub

Friend Function GetCompareMode() As PD_FloodCompare
    GetCompareMode = m_CompareMode
End Function

Friend Sub SetCompareMode(ByVal newCompareMode As PD_FloodCompare)
    m_CompareMode = newCompareMode
End Sub

Friend Function GetInitialPoint() As PointAPI
    GetInitialPoint = m_InitPoint
End Function

Friend Sub SetInitialPoint(ByVal startX As Long, ByVal startY As Long)
    m_InitPoint.x = startX
    m_InitPoint.y = startY
End Sub

Friend Function GetSearchMode() As PD_FloodSearch
    GetSearchMode = m_SearchMode
End Function

Friend Sub SetSearchMode(ByVal newSearchMode As PD_FloodSearch)
    m_SearchMode = newSearchMode
End Sub

'Tolerance is returned on a [0, 100] scale.
Friend Function GetTolerance() As Double
    GetTolerance = m_Tolerance
End Function

'Tolerance is set on a [0, 100] scale.
Friend Sub SetTolerance(ByVal newTolerance As Double)
    m_Tolerance = newTolerance
    If (m_Tolerance > 100#) Then m_Tolerance = 100#
    If (m_Tolerance < 0#) Then m_Tolerance = 0#
End Sub

'Initiate a flood fill operation.  This class doesn't actually fill anything; what it does is fill a 32-bpp destination DIB
' (if supplied - you CAN pass Nothing if you only want a path) with an alpha map of the flood results, where
' black/transparent = unfilled, white/opaque = filled, gray/partially transparent = partially filled.
'
'This approach lets the caller use the flood results however they want, without having to modify this class to match.
'
'In a change from earlier versions, the caller must also now provide a target pdPath object.  This path is filled with a
' vector representation of the filled area, and it is also used to produce antialiased fill results.  The caller can obviously
' discard this path if they don't need it, but it *will* be used regardless.
Friend Function InitiateFloodFill(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByRef dstOutlinePath As pd2DPath) As Boolean
    
    'Make sure the passed x/y coords are valid.  If they aren't, exit now.
    If (m_InitPoint.x < 0) Or (m_InitPoint.y < 0) Or (m_InitPoint.x >= srcDIB.GetDIBWidth) Or (m_InitPoint.y >= srcDIB.GetDIBHeight) Then
        Debug.Print "Invalid flood fill location requested.  Abandoning flood fill now.  (" & m_InitPoint.x & ", " & m_InitPoint.y & ")"
        InitiateFloodFill = False
        Exit Function
    End If
    
    'Flood filling can be expensive.  We may want to time it in debug builds.
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'If a destination DIB was supplied, initialize it now
    If (Not dstDIB Is Nothing) Then
        If (dstDIB.GetDIBWidth = srcDIB.GetDIBWidth) And (dstDIB.GetDIBHeight = srcDIB.GetDIBHeight) And (dstDIB.GetDIBColorDepth = 32) Then
            dstDIB.ResetDIB 0
        Else
            dstDIB.CreateBlank srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 32, 0
        End If
    End If
    
    'Prep a tracking array
    Dim xBound As Long, yBound As Long
    xBound = srcDIB.GetDIBWidth - 1
    yBound = srcDIB.GetDIBHeight - 1
    
    If (m_BoundsX <> xBound) Or (m_BoundsY <> yBound) Then
        ReDim m_AlreadyChecked(0 To xBound, 0 To yBound) As Byte
        ReDim m_FillResults(0 To xBound, 0 To yBound) As Byte
        m_BoundsX = xBound
        m_BoundsY = yBound
    Else
        FillMemory VarPtr(m_AlreadyChecked(0, 0)), (xBound + 1) * (yBound + 1), 0
        FillMemory VarPtr(m_FillResults(0, 0)), (xBound + 1) * (yBound + 1), 0
    End If
    
    'Based on the specified search mode, call the appropriate flood function
    If (m_SearchMode = pdfs_Contiguous) Then
        
        'Contiguous floodfills use a new scanline-based implementation.  This provides a 30-40% performance improvement
        ' over a naive implementation, although the code is quite a bit more complex.
        FloodFillContiguous_Scanline srcDIB
        
    ElseIf (m_SearchMode = pdfs_WholeImage) Then
        FloodFillGlobal srcDIB
    End If
    
    'pdDebug.LogAction "FF-MAP: " & Format$(VBHacks.GetTimerDifferenceNow(startTime) * 1000, "#0") & " ms"
    VBHacks.GetHighResTime startTime
    
    'From the flood fill results, generate an outline path
    GenerateFloodFillOutline dstOutlinePath
    
    'pdDebug.LogAction "FF-OUTLINE: " & Format$(VBHacks.GetTimerDifferenceNow(startTime) * 1000, "#0") & " ms"
    VBHacks.GetHighResTime startTime
    
    'Finally, render our floodfill results onto the destination DIB, if one was provided
    If (Not dstDIB Is Nothing) Then
        
        Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDC cSurface, dstDIB.GetDIBDC, m_AntialiasingMode
        If m_AntialiasingMode Then
            cSurface.SetSurfacePixelOffset P2_PO_Half
            PDMath.SimplifyPathFromMarchingSquares dstOutlinePath
        Else
            cSurface.SetSurfacePixelOffset P2_PO_Normal
        End If
        
        If (m_Brush Is Nothing) Then Drawing2D.QuickCreateSolidBrush m_Brush, RGB(255, 255, 255), 100!
        
        PD2D.FillPath cSurface, m_Brush, dstOutlinePath
        Set cSurface = Nothing
        
        'pdDebug.LogAction "FF-RENDER: " & Format$(VBHacks.GetTimerDifferenceNow(startTime) * 1000, "#0") & " ms"
        
    End If
    
    InitiateFloodFill = True

End Function

'*AFTER* a contiguous floodfill has been calculated, you can call this function to generate a vector representation of the
' filled area's outline.  Note that any existing path data is *not* erased; the path is simply appended to existing data,
' if any exists.
Private Sub GenerateFloodFillOutline(ByRef dstPath As pd2DPath)
    
    'PD's current outline detector requires an empty border around the outside of the target area.  Generate such an array now.
    Dim xBound As Long, yBound As Long
    xBound = m_BoundsX + 2
    yBound = m_BoundsY + 2

    If (m_OutlineBoundsX <> xBound) Or (m_OutlineBoundsY <> yBound) Then
        ReDim m_OutlineCopy(0 To xBound, 0 To yBound) As Byte
        m_OutlineBoundsX = xBound
        m_OutlineBoundsY = yBound
    
    'Normally, we would wipe the outline array prior to working with it, but because we're just gonna
    ' fill the bytes with a manual memcpy loop, we can safely ignore a prior fill command.
    ' (The outer borders will always be left blank, by design.)
    'Else
        'FillMemory VarPtr(m_OutlineCopy(0, 0)), (xBound + 1) * (yBound + 1), 0
    End If

    'We need to copy all lines from the boundary check array to our outline array, offsetting them by
    ' (1) in each direction.  This guarantees a boundary of zeroes around the target image.

    'Because VB arrays are row-major, we're going to copy contiguous rows.
    Dim copySize As Long
    copySize = m_BoundsX + 1
    
    Dim y As Long
    For y = 0 To m_BoundsY
        CopyMemoryStrict VarPtr(m_OutlineCopy(0, y + 1)) + 1, VarPtr(m_FillResults(0, y)), copySize
    Next y
    
    'The m_OutlineCopy array now contains a valid copy of the filled area, with guaranteed blank boundary lines
    
    'Initiate the outline search.  (Note that we can safely start from position (1, 1), because we have already inserted blank lines
    ' around the exterior of the fill map.)
    If (m_FillOutline Is Nothing) Then Set m_FillOutline = New pdEdgeDetector
    m_FillOutline.FindAllEdges dstPath, m_OutlineCopy, 1, 1, m_BoundsX + 1, m_BoundsY + 1, -1, -1
    
    'The edge detector has already filled the destination path, so our work here is done!
    
End Sub

'Perform a contiguous (default) flood fill, using horizontal scanlines to improve performance.
'
'IMPORTANT NOTE!  As of v7.0, the source DIB is required to be 32-bpp.  Passing a 24-bpp image will cause a hard crash.
' (This matches PD's internal conversion to always-enforced 32-bpp sources.)
Private Function FloodFillContiguous_Scanline(ByRef srcDIB As pdDIB) As Boolean

    'Reset the stack.  Note that we don't actually resize the stack; this is an optimization technique to improve performance
    ' if this class is used multiple times in a row.
    m_StackPosition = -1
    
    'Predetermine upper bounds for x/y checks
    Dim xBound As Long, yBound As Long
    xBound = srcDIB.GetDIBWidth - 1
    yBound = srcDIB.GetDIBHeight - 1
    
    'Populate the initial stack point
    PushOntoStack m_InitPoint.x, m_InitPoint.y
    
    Dim x As Long, y As Long, xStride As Long
    
    'Generate direct references to the source and destination DIB data
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    'A number of local variables are used to help optimize the flood function
    Dim isWithinTolerance As Boolean
    Dim modifiedTolerance As Long
    
    'Populate our reference comparison values
    Dim r As Long, g As Long, b As Long, a As Long, l As Long
    Dim refR As Long, refG As Long, refB As Long, refA As Long, refL As Long
    
    Dim thisValue As Long
    
    xStride = m_InitPoint.x * 4
    y = m_InitPoint.y
    
    refB = srcImageData(xStride, y)
    refG = srcImageData(xStride + 1, y)
    refR = srcImageData(xStride + 2, y)
    refA = srcImageData(xStride + 3, y)
    
    refL = 213 * refR + 715 * refG + 72 * refB
    
    'Calculate a reference tolerance value, which serves as the base for the flood fill
    Select Case m_CompareMode
    
        Case pdfc_Color
            modifiedTolerance = m_Tolerance * 2.55
            modifiedTolerance = (modifiedTolerance * modifiedTolerance) * 3
        
        'Composite results do not require a base value.  They are independently processed against the reference
        ' RGB values as we go.  However, to accelerate the required check, we premultiply the requested tolerance
        ' by 4, to avoid the need for a divide function in the inner loop
        Case pdfc_Composite
            modifiedTolerance = m_Tolerance * 2.55
            modifiedTolerance = (modifiedTolerance * modifiedTolerance) * 4
            
        Case pdfc_Luminance
            modifiedTolerance = m_Tolerance * 2550
            
        Case pdfc_Red
            modifiedTolerance = m_Tolerance * 2.55
            
        Case pdfc_Green
            modifiedTolerance = m_Tolerance * 2.55
            
        Case pdfc_Blue
            modifiedTolerance = m_Tolerance * 2.55
            
        Case pdfc_Alpha
            modifiedTolerance = m_Tolerance * 2.55
    
    End Select
    
    Dim scanRight As Boolean, scanLeft As Boolean
    Dim leftBound As Long, rightBound As Long, i As Long, yIn As Long
    
    'To improve performance, we're going to point transient 1D arrays at certain 2D array rows during processing.
    ' This requires unsafe array manipulation, but it can be significantly faster than 2D array accesses (and because
    ' this function is scanline-oriented, the gains are even more significant).
    Dim tmpImageLine() As Byte, tmpImageSA As SafeArray1D
    
    'Populate the safearray struct's unchanging values
    Dim srcImageBasePointer As Long, srcImageStride As Long
    srcImageBasePointer = srcDIB.GetDIBPointer
    srcImageStride = srcDIB.GetDIBStride
    
    With tmpImageSA
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcImageStride
        
        'pvData *will* change as the function goes along, but let's at least start with a safe value
        .pvData = srcImageBasePointer
        
    End With
    
    'Point the uninitialized temporary array at our custom-built SafeArray struct
    PutMem4 VarPtrArray(tmpImageLine()), VarPtr(tmpImageSA)
    
    'Repeat the above steps, but for a 1D array that points at the "already checked" tracking array
    Dim tmpTrackingLine() As Byte, tmpTrackingSA As SafeArray1D
    Dim srcTrackingBasePointer As Long, srcTrackingStride As Long
    srcTrackingBasePointer = VarPtr(m_AlreadyChecked(0, 0))
    srcTrackingStride = m_BoundsX + 1
    
    With tmpTrackingSA
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcTrackingStride
        .pvData = srcTrackingBasePointer
    End With
    
    PutMem4 VarPtrArray(tmpTrackingLine()), VarPtr(tmpTrackingSA)
    
    'And finally, repeat the above steps, but for a 1D array that points at the "fill results" tracking array
    Dim tmpFillLine() As Byte, tmpFillSA As SafeArray1D
    Dim srcFillBasePointer As Long, srcFillStride As Long
    srcFillBasePointer = VarPtr(m_FillResults(0, 0))
    srcFillStride = m_BoundsX + 1
    
    With tmpFillSA
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcFillStride
        .pvData = srcFillBasePointer
    End With
    
    PutMem4 VarPtrArray(tmpFillLine()), VarPtr(tmpFillSA)
    
    'Start processing the stack!
    Do
    
        'Reset the tolerance check
        isWithinTolerance = False
        
        'Retrieve the next point from the stack.  Normally we would do this with a call to the pop function, e.g.:
        'PopFromStack x, y
        '
        '...but it's faster to inline the function like so:
        x = m_Stack(m_StackPosition).x
        y = m_Stack(m_StackPosition).y
        m_StackPosition = m_StackPosition - 1
        
        'Retrieve RGB/A values for this point
        tmpImageSA.pvData = srcImageBasePointer + (y * srcImageStride)
        xStride = x * 4
        b = tmpImageLine(xStride)
        g = tmpImageLine(xStride + 1)
        r = tmpImageLine(xStride + 2)
        a = tmpImageLine(xStride + 3)
        
        'Compare this pixel against the reference
        Select Case m_CompareMode
            Case pdfc_Color
                r = (r - refR)
                g = (g - refG)
                b = (b - refB)
                thisValue = r * r + g * g + b * b
                isWithinTolerance = (thisValue <= modifiedTolerance)
            Case pdfc_Composite
                r = (r - refR)
                g = (g - refG)
                b = (b - refB)
                a = (a - refA)
                thisValue = r * r + g * g + b * b + a * a
                isWithinTolerance = (thisValue <= modifiedTolerance)
            Case pdfc_Luminance
                l = 213 * r + 715 * g + 72 * b
                isWithinTolerance = (Abs(l - refL) <= modifiedTolerance)
            Case pdfc_Red
                isWithinTolerance = (Abs(r - refR) <= modifiedTolerance)
            Case pdfc_Green
                isWithinTolerance = (Abs(g - refG) <= modifiedTolerance)
            Case pdfc_Blue
                isWithinTolerance = (Abs(b - refB) <= modifiedTolerance)
            Case pdfc_Alpha
                isWithinTolerance = (Abs(a - refA) <= modifiedTolerance)
        End Select
        
        'If this value is within the requested tolerance, mark it on the destination map
        If isWithinTolerance Then
            
            'Mark this pixel as filled
            tmpTrackingSA.pvData = srcTrackingBasePointer + (y * srcTrackingStride)
            tmpTrackingLine(x) = 2
            
            tmpFillSA.pvData = srcFillBasePointer + (y * srcFillStride)
            tmpFillLine(x) = 255
            
            'Next, we're going to do a full scanline check in both the left and right directions.  Start with the
            ' left direction.
            leftBound = x - 1
            scanLeft = (leftBound >= 0)
            Do While scanLeft
                If (tmpTrackingLine(leftBound) = 0) Then
                    
                    'Retrieve RGB/A values for this point
                    xStride = leftBound * 4
                    b = tmpImageLine(xStride)
                    g = tmpImageLine(xStride + 1)
                    r = tmpImageLine(xStride + 2)
                    a = tmpImageLine(xStride + 3)
                
                    'Compare this pixel against the reference
                    Select Case m_CompareMode
                        Case pdfc_Color
                            r = (r - refR)
                            g = (g - refG)
                            b = (b - refB)
                            thisValue = r * r + g * g + b * b
                            isWithinTolerance = (thisValue <= modifiedTolerance)
                        Case pdfc_Composite
                            r = (r - refR)
                            g = (g - refG)
                            b = (b - refB)
                            a = (a - refA)
                            thisValue = r * r + g * g + b * b + a * a
                            isWithinTolerance = (thisValue <= modifiedTolerance)
                        Case pdfc_Luminance
                            l = 213 * r + 715 * g + 72 * b
                            isWithinTolerance = (Abs(l - refL) <= modifiedTolerance)
                        Case pdfc_Red
                            isWithinTolerance = (Abs(r - refR) <= modifiedTolerance)
                        Case pdfc_Green
                            isWithinTolerance = (Abs(g - refG) <= modifiedTolerance)
                        Case pdfc_Blue
                            isWithinTolerance = (Abs(b - refB) <= modifiedTolerance)
                        Case pdfc_Alpha
                            isWithinTolerance = (Abs(a - refA) <= modifiedTolerance)
                    End Select
                    
                    'If this value is within the requested tolerance, mark it on the destination map
                    If isWithinTolerance Then
                        tmpTrackingLine(leftBound) = 2
                        tmpFillLine(leftBound) = 255
                        leftBound = leftBound - 1
                        scanLeft = (leftBound >= 0)
                    Else
                        tmpTrackingLine(leftBound) = 1
                        scanLeft = False
                    End If
                    
                Else
                    scanLeft = False
                End If
            Loop
            
            leftBound = leftBound + 1
            
            rightBound = x + 1
            scanRight = (rightBound <= xBound)
            Do While scanRight
                If (tmpTrackingLine(rightBound) = 0) Then
                    
                    'Retrieve RGB/A values for this point
                    xStride = rightBound * 4
                    b = tmpImageLine(xStride)
                    g = tmpImageLine(xStride + 1)
                    r = tmpImageLine(xStride + 2)
                    a = tmpImageLine(xStride + 3)
                
                    'Compare this pixel against the reference
                    Select Case m_CompareMode
                        Case pdfc_Color
                            r = (r - refR)
                            g = (g - refG)
                            b = (b - refB)
                            thisValue = r * r + g * g + b * b
                            isWithinTolerance = (thisValue <= modifiedTolerance)
                        Case pdfc_Composite
                            r = (r - refR)
                            g = (g - refG)
                            b = (b - refB)
                            a = (a - refA)
                            thisValue = r * r + g * g + b * b + a * a
                            isWithinTolerance = (thisValue <= modifiedTolerance)
                        Case pdfc_Luminance
                            l = 213 * r + 715 * g + 72 * b
                            isWithinTolerance = (Abs(l - refL) <= modifiedTolerance)
                        Case pdfc_Red
                            isWithinTolerance = (Abs(r - refR) <= modifiedTolerance)
                        Case pdfc_Green
                            isWithinTolerance = (Abs(g - refG) <= modifiedTolerance)
                        Case pdfc_Blue
                            isWithinTolerance = (Abs(b - refB) <= modifiedTolerance)
                        Case pdfc_Alpha
                            isWithinTolerance = (Abs(a - refA) <= modifiedTolerance)
                    End Select
                    
                    'If this value is within the requested tolerance, mark it on the destination map
                    If isWithinTolerance Then
                        tmpTrackingLine(rightBound) = 2
                        tmpFillLine(rightBound) = 255
                        rightBound = rightBound + 1
                        scanRight = (rightBound <= xBound)
                    Else
                        tmpTrackingLine(rightBound) = 1
                        scanRight = False
                    End If
                    
                Else
                    scanRight = False
                End If
            Loop
            
            rightBound = rightBound - 1
            
            'Finally, push all neighboring pixels onto the stack.  Normally we would do this via the cleaner "PushOntoStack" function,
            ' but for performance reasons, we inline the stack requests.
            
            'Start by figuring out the maximum stack size we may need (if every neighboring point would be added), and ensuring
            ' at least that much space is available.
            yIn = m_StackPosition + (rightBound - leftBound + 1) * 2
            If (m_StackHeight < yIn) Then
                m_StackHeight = yIn * 2 + 1
                ReDim Preserve m_Stack(0 To m_StackHeight) As PointAPI
            End If
            
            'Now that a safe stack size is guaranteed, we can push the pixels above and below this one onto the stack
            ' in one fell swoop, without worrying about safe memory allocations.
            If (y > 0) Then
                yIn = y - 1
                tmpTrackingSA.pvData = srcTrackingBasePointer + (yIn * srcTrackingStride)
                For i = leftBound To rightBound
                    If (tmpTrackingLine(i) = 0) Then
                        tmpTrackingLine(i) = 1
                        m_StackPosition = m_StackPosition + 1
                        m_Stack(m_StackPosition).x = i
                        m_Stack(m_StackPosition).y = yIn
                    End If
                Next i
            End If
            
            If (y < yBound) Then
                yIn = y + 1
                tmpTrackingSA.pvData = srcTrackingBasePointer + (yIn * srcTrackingStride)
                For i = leftBound To rightBound
                    If (tmpTrackingLine(i) = 0) Then
                        tmpTrackingLine(i) = 1
                        m_StackPosition = m_StackPosition + 1
                        m_Stack(m_StackPosition).x = i
                        m_Stack(m_StackPosition).y = yIn
                    End If
                Next i
            End If
            
        End If
                
    'As long as there are more stack points to process, rinse and repeat
    Loop While (m_StackPosition >= 0)
    
    'Release our array references
    srcDIB.UnwrapArrayFromDIB srcImageData
    PutMem4 VarPtrArray(tmpImageLine()), 0&
    PutMem4 VarPtrArray(tmpTrackingLine()), 0&
    PutMem4 VarPtrArray(tmpFillLine()), 0&
    
    FloodFillContiguous_Scanline = True
    
End Function

'Perform a full-image, non-contiguous flood fill.
'
'IMPORTANT NOTE!  As of v7.0, the source DIB is required to be 32-bpp.  Passing a 24-bpp image will cause a hard crash.
' (This matches PD's internal conversion to always-enforced 32-bpp sources.)
Private Function FloodFillGlobal(ByRef srcDIB As pdDIB) As Boolean
    
    'Predetermine upper bounds for x/y checks
    Dim xBound As Long, yBound As Long
    xBound = srcDIB.GetDIBWidth - 1
    yBound = srcDIB.GetDIBHeight - 1
    
    'Make sure 24 and 32bpp sources are both handled correctly
    Dim x As Long, y As Long, xStride As Long
    
    'Generate direct references to the source and destination DIB data
    Dim srcImageData() As Byte, srcSA As SafeArray1D
    srcDIB.WrapArrayAroundScanline srcImageData, srcSA, 0
    
    Dim dibPtr As Long, dibStride As Long
    dibPtr = srcSA.pvData
    dibStride = srcSA.cElements
    
    'A number of local variables are used to help optimize the flood function
    Dim isWithinTolerance As Boolean, modifiedTolerance As Double
    
    'Populate our reference comparison values
    Dim r As Long, g As Long, b As Long, a As Long, l As Long
    Dim refR As Long, refG As Long, refB As Long, refA As Long, refL As Long
    Dim thisValue As Double
    
    x = m_InitPoint.x * 4
    y = m_InitPoint.y
    srcSA.pvData = dibPtr + (y * dibStride)
    
    refB = srcImageData(x)
    refG = srcImageData(x + 1)
    refR = srcImageData(x + 2)
    refA = srcImageData(x + 3)
    
    refL = (213 * refR + 715 * refG + 72 * refB)
    
    'Calculate a reference tolerance value, which serves as the base for the flood fill
    If (m_CompareMode = pdfc_Composite) Then
        'Composite results do not require a base value, as they are independently processed against the reference
        ' RGBA values as we go.  However, to accelerate the required check, we premultiply the requested tolerance
        ' by 4, to avoid the need for a divide function in the inner loop
        modifiedTolerance = m_Tolerance * 2.55
        modifiedTolerance = (modifiedTolerance * modifiedTolerance) * 3
            
    ElseIf (m_CompareMode = pdfc_Color) Then
        modifiedTolerance = m_Tolerance * 2.55
        modifiedTolerance = (modifiedTolerance * modifiedTolerance) * 4
        
    ElseIf (m_CompareMode = pdfc_Luminance) Then
        modifiedTolerance = m_Tolerance * 2550
            
    ElseIf (m_CompareMode = pdfc_Red) Then
        modifiedTolerance = m_Tolerance * 2.55
            
    ElseIf (m_CompareMode = pdfc_Green) Then
        modifiedTolerance = m_Tolerance * 2.55
            
    ElseIf (m_CompareMode = pdfc_Blue) Then
        modifiedTolerance = m_Tolerance * 2.55
            
    ElseIf (m_CompareMode = pdfc_Alpha) Then
        modifiedTolerance = m_Tolerance * 2.55
    
    End If
    
    'Start processing the image!
    For y = 0 To yBound
        srcSA.pvData = dibPtr + (y * dibStride)
    For x = 0 To xBound
        
        'Reset the tolerance check
        isWithinTolerance = False
        
        'Retrieve RGB/A values for this point
        xStride = x * 4
        b = srcImageData(xStride)
        g = srcImageData(xStride + 1)
        r = srcImageData(xStride + 2)
        a = srcImageData(xStride + 3)
    
        'Compare this pixel against the reference
        Select Case m_CompareMode
        
            Case pdfc_Color
                r = (r - refR)
                g = (g - refG)
                b = (b - refB)
                thisValue = r * r + g * g + b * b
                isWithinTolerance = (thisValue <= modifiedTolerance)
                
            Case pdfc_Composite
                r = (r - refR)
                g = (g - refG)
                b = (b - refB)
                a = (a - refA)
                thisValue = r * r + g * g + b * b + a * a
                isWithinTolerance = (thisValue <= modifiedTolerance)
            
            Case pdfc_Luminance
                l = 213 * r + 715 * g + 72 * b
                isWithinTolerance = (Abs(l - refL) <= modifiedTolerance)
            
            Case pdfc_Red
                isWithinTolerance = (Abs(r - refR) <= modifiedTolerance)
            
            Case pdfc_Green
                isWithinTolerance = (Abs(g - refG) <= modifiedTolerance)
                
            Case pdfc_Blue
                isWithinTolerance = (Abs(b - refB) <= modifiedTolerance)
            
            Case pdfc_Alpha
                isWithinTolerance = (Abs(a - refA) <= modifiedTolerance)
        
        End Select
        
        'If this value is within the requested tolerance, mark it on the destination map
        If isWithinTolerance Then m_FillResults(x, y) = 255
        
    Next x
    Next y
    
    'Release our array references
    srcDIB.UnwrapArrayFromDIB srcImageData
    
    FloodFillGlobal = True
    
End Function

'Stack helper functions
Private Sub PushOntoStack(ByVal x As Long, ByVal y As Long)
    
    m_StackPosition = m_StackPosition + 1
    
    'Resize the stack as necessary
    If (m_StackPosition > m_StackHeight) Then
        m_StackHeight = m_StackHeight * 2 + 1
        ReDim Preserve m_Stack(0 To m_StackHeight) As PointAPI
    End If
    
    'Mark this point as "due to be checked", so it does not get re-checked
    m_AlreadyChecked(x, y) = 1
    
    'Add the point to the stack
    m_Stack(m_StackPosition).x = x
    m_Stack(m_StackPosition).y = y
    
End Sub

'This class automatically handles its own memory management, but if you're using a persistent instance of it, you can manually call
' this function to free up cached resources.  (This class doesn't automatically free things like its custom flood fill stack, to make
' subsequent flood fill requests faster.)
Friend Sub FreeUpResources()
    
    m_StackHeight = INITIAL_STACK_HEIGHT - 1
    ReDim m_Stack(0 To m_StackHeight) As PointAPI
    
    m_BoundsX = 0
    m_BoundsY = 0
    ReDim m_AlreadyChecked(0, 0) As Byte
    
    m_OutlineBoundsX = 0
    m_OutlineBoundsY = 0
    ReDim m_FillResults(0, 0) As Byte
    
    Set m_Brush = Nothing
    Set m_FillOutline = Nothing
    
End Sub

Private Sub Class_Initialize()

    'Reset all stack values
    m_StackPosition = 0
    m_StackHeight = INITIAL_STACK_HEIGHT - 1
    ReDim m_Stack(0 To m_StackHeight) As PointAPI
    
    'Reset our check and fill arrays
    m_BoundsX = 0
    m_BoundsY = 0
    ReDim m_AlreadyChecked(0, 0) As Byte
    
    m_OutlineBoundsX = 0
    m_OutlineBoundsY = 0
    ReDim m_FillResults(0, 0) As Byte
    
    'Composite is the default tolerance mode
    m_CompareMode = pdfc_Composite
    
End Sub
