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
'Copyright 2014-2017 by Tanner Helland
'Created: 11/October/14
'Last updated: 26/February/17
'Last update: various performance improvements to all floodfill types
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'A stack is used to track pixels that need to be checked
Private m_Stack() As POINTAPI
Private m_StackPosition As Long
Private m_StackHeight As Long
Private Const INITIAL_STACK_HEIGHT As Long = 4096

'To reduce iterations, this function tracks pixels that have already been added to the stack.  The boundary values are
' 0-based, so for a 10x10 image, these values will be (9, 9), respectively.
Private m_BoundsX As Long, m_BoundsY As Long
Private m_AlreadyChecked() As Byte

'If outlines are requested, this array will be used to store outline data.
Private m_OutlineBoundsX As Long, m_OutlineBoundsY As Long
Private m_OutlineCopy() As Byte

'Tolerance allows the user to control the strength of the flood
Private m_Tolerance As Double

'Different compare modes can be used to obtain better results.
Public Enum PD_FLOOD_COMPARE
    PDFC_COMPOSITE = 0
    PDFC_COLOR = 1
    PDFC_LUMINANCE = 2
    PDFC_RED = 3
    PDFC_GREEN = 4
    PDFC_BLUE = 5
    PDFC_ALPHA = 6
End Enum

#If False Then
    Private Const PDFC_COMPOSITE = 0, PDFC_COLOR = 1, PDFC_LUMINANCE = 2, PDFC_RED = 3, PDFC_GREEN = 4, PDFC_BLUE = 5, PDFC_ALPHA = 6
#End If

Private m_CompareMode As PD_FLOOD_COMPARE

'Two search methods are supported: contiguous region based on the initial point (default behavior), or the full image
Public Enum PD_FLOOD_SEARCH
    PDFS_CONTIGUOUS = 0
    PDFS_WHOLEIMAGE = 1
End Enum

#If False Then
    Private Const PDFS_CONTIGUOUS = 0, PDFS_WHOLEIMAGE = 1
#End If

Private m_SearchMode As PD_FLOOD_SEARCH

'A custom antialiasing technique can be used to soften the floodfill results
Private m_AntialiasingMode As Boolean

'Obviously the function needs a starting x/y position
Private m_InitPoint As POINTAPI

'Fast memory ops
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal dstPointer As Long, ByVal srcPointer As Long, ByVal numOfBytes As Long)


'Get/set functions for all relevant flood fill parameters
Public Function GetAntialiasingMode() As Boolean
    GetAntialiasingMode = m_AntialiasingMode
End Function

Public Sub SetAntialiasingMode(ByVal newAntialiasingMode As Boolean)
    m_AntialiasingMode = newAntialiasingMode
End Sub

Public Function GetCompareMode() As PD_FLOOD_COMPARE
    GetCompareMode = m_CompareMode
End Function

Public Function SetCompareMode(ByVal newCompareMode As PD_FLOOD_COMPARE)
    m_CompareMode = newCompareMode
End Function

Friend Function GetInitialPoint() As POINTAPI
    GetInitialPoint = m_InitPoint
End Function

Public Sub SetInitialPoint(ByVal startX As Long, ByVal startY As Long)
    m_InitPoint.x = startX
    m_InitPoint.y = startY
End Sub

Public Function GetSearchMode() As PD_FLOOD_SEARCH
    GetSearchMode = m_SearchMode
End Function

Public Sub SetSearchMode(ByVal newSearchMode As PD_FLOOD_SEARCH)
    m_SearchMode = newSearchMode
End Sub

Public Function GetTolerance() As Double
    GetTolerance = m_Tolerance
End Function

Public Sub SetTolerance(ByVal newTolerance As Double)
    m_Tolerance = newTolerance
End Sub

'Initiate a flood fill operation.  This class doesn't actually fill anything; what it does is fill a 32-bpp destination DIB
' with an alpha of the flood results (black/transparent = unfilled, white/opaque = filled, gray/partially transparent = partially filled).
' This allows the caller to use the flood results however they want, without having to modify this class to match.
'
'Optionally, the caller can also provide a target pdPath object; if supplied, this path will be filled with a vector representation
' of the filled area.  (IMPORTANT DISCLAIMER: this feature currently only works with contiguous fills, not global ones.)
Public Function InitiateFloodFill(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByRef dstOutlinePath As pd2DPath = Nothing) As Boolean
    
    'Make sure the passed x/y coords are valid.  If they aren't, exit now.
    If (m_InitPoint.x < 0) Or (m_InitPoint.y < 0) Or (m_InitPoint.x >= srcDIB.GetDIBWidth) Or (m_InitPoint.y >= srcDIB.GetDIBHeight) Then
        Debug.Print "Invalid flood fill location requested.  Abandoning flood fill now."
        InitiateFloodFill = False
        Exit Function
    End If
    
    'Initialize the destination DIB
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    If (dstDIB.GetDIBWidth = srcDIB.GetDIBWidth) And (dstDIB.GetDIBHeight = srcDIB.GetDIBHeight) And (dstDIB.GetDIBColorDepth = 32) Then
        dstDIB.ResetDIB 0
    Else
        dstDIB.CreateBlank srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 32, 0
    End If
    
    'Based on the specified search mode, call the appropriate flood function
    If (m_SearchMode = PDFS_CONTIGUOUS) Then
        FloodFillContiguous srcDIB, dstDIB
        
        'If the caller wants us to return an outline path, populate that here
        If (Not dstOutlinePath Is Nothing) Then GenerateOutlinePath dstOutlinePath
        
    ElseIf (m_SearchMode = PDFS_WHOLEIMAGE) Then
        FloodFillGlobal srcDIB, dstDIB
    End If
    
    'If antialiasing was requested, apply it now.
    If m_AntialiasingMode Then
        
        'For contiguous fill modes, we use a custom AA solution that provides better results.
        If (m_SearchMode = PDFS_CONTIGUOUS) Then
            FloodFillCustomAA dstDIB
        
        'For global fill modes, a simple quick blur will do
        Else
            Filters_Layers.QuickBlurDIB dstDIB, 1
        End If
        
    End If
    
    InitiateFloodFill = True

End Function

'*AFTER* a floodfill has been calculated, you can call this function to generate a vector representation of the filled area's outline.
' Note that any existing path data is *not* erased; the path is simply appended to existing data, if any exists.
Private Sub GenerateOutlinePath(ByRef dstPath As pd2DPath)

    'PD's current outline detector requires an empty border around the outside of the target area.  Generate such an array now.
    Dim xBound As Long, yBound As Long
    xBound = m_BoundsX + 2
    yBound = m_BoundsY + 2
    
    If (m_OutlineBoundsX <> xBound) Or (m_OutlineBoundsY <> yBound) Then
        ReDim m_OutlineCopy(0 To xBound, 0 To yBound) As Byte
        m_OutlineBoundsX = xBound
        m_OutlineBoundsY = yBound
    Else
        FillMemory VarPtr(m_OutlineCopy(0, 0)), (xBound + 1) * (yBound + 1), 0
    End If
    
    'We need to copy all lines from the boundary check array to our outline array, offsetting them by (1) in each direction.
    ' This guarantees a boundary of zeroes around the target image.
    
    'Because VB arrays are row-major, we're going to copy contiguous rows.
    Dim copySize As Long
    copySize = m_BoundsX
    
    Dim x As Long, y As Long
    For y = 0 To m_BoundsY
        CopyMemory_Strict VarPtr(m_OutlineCopy(0, y + 1)) + 1, VarPtr(m_AlreadyChecked(0, y)), m_BoundsX
    Next y
    
    'The m_OutlineCopy array now contains a valid copy of the filled area, with guaranteed blank boundary lines
    
    'Prep the target path and the outline enumerator
    If (dstPath Is Nothing) Then Set dstPath = New pd2DPath
    
    Dim cOutline As pdEdgeDetector
    Set cOutline = New pdEdgeDetector
    
    'Find a starting point
    Dim startX As Long, startY As Long
    cOutline.FindStartingPoint m_OutlineCopy, 1, 1, m_BoundsX + 1, m_BoundsY + 1, startX, startY
    
    'Initiate the outline search
    cOutline.FindEdges m_OutlineCopy, startX, startY, -1, -1
    
    'Retrieve the final point list
    cOutline.RetrieveFinalPolygonIntoPath dstPath, True
    
End Sub

'Perform a contiguous (default) flood fill.  Note that dstDIB is guaranteed to be properly initialized.
'
'IMPORTANT NOTE!  As of v7.0, both source and destination DIBs are required to be 32-bpp.  Passing a 24-bpp image will
' cause a hard crash.  (This matches PD's internal conversion to always-enforced 32-bpp sources.)
Private Function FloodFillContiguous(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Boolean

    'Reset the stack.  Note that we don't actually resize the stack; this is an optimization technique to improve performance
    ' if this class is used multiple times in a row.
    m_StackPosition = -1
    
    'Predetermine upper bounds for x/y checks
    Dim xBound As Long, yBound As Long
    xBound = srcDIB.GetDIBWidth - 1
    yBound = srcDIB.GetDIBHeight - 1
    
    'Reset the "has been filled" array
    If (m_BoundsX <> xBound) Or (m_BoundsY <> yBound) Then
        ReDim m_AlreadyChecked(0 To xBound, 0 To yBound) As Byte
        m_BoundsX = xBound
        m_BoundsY = yBound
    Else
        FillMemory VarPtr(m_AlreadyChecked(0, 0)), (xBound + 1) * (yBound + 1), 0
    End If
    
    'Populate the initial stack point
    PushOntoStack m_InitPoint.x, m_InitPoint.y
    
    Dim x As Long, y As Long, quickX As Long, dstQuickX As Long
    
    'Generate direct references to the source and destination DIB data
    Dim srcImageData() As Byte, dstImageData() As Long
    Dim srcSA As SAFEARRAY2D, dstSA As SAFEARRAY2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'For performance reasons, the destination DIB uses a Long-type array instead of a byte array (as we don't need access to individual bytes)
    dstDIB.WrapLongArrayAroundDIB dstImageData, dstSA
    
    'A number of local variables are used to help optimize the flood function
    Dim isWithinTolerance As Boolean
    Dim modifiedTolerance As Double
    
    'Populate our reference comparison values
    Dim r As Long, g As Long, b As Long, a As Long, l As Long
    Dim refR As Long, refG As Long, refB As Long, refA As Long, refL As Long
    
    Dim thisValue As Double
    
    quickX = m_InitPoint.x * 4
    y = m_InitPoint.y
    
    refB = srcImageData(quickX, y)
    refG = srcImageData(quickX + 1, y)
    refR = srcImageData(quickX + 2, y)
    refA = srcImageData(quickX + 3, y)
    
    refL = 213 * refR + 715 * refG + 72 * refB
    
    'Calculate a reference tolerance value, which serves as the base for the flood fill
    Select Case m_CompareMode
    
        Case PDFC_COMPOSITE
            'Composite results do not require a base value, as they are independently processed against the reference
            ' RGB values as we go.  However, to accelerate the required check, we premultiply the requested tolerance
            ' by 4, to avoid the need for a divide function in the inner loop
            modifiedTolerance = m_Tolerance * 4
            
        Case PDFC_COLOR
            modifiedTolerance = m_Tolerance * 3
        
        Case PDFC_LUMINANCE
        
            'To save time on the inner loop, we don't divide luminance by 1000; to make this work, we must change the
            ' tolerance range to [0, 1000] instead of [0, 255.0]
            modifiedTolerance = m_Tolerance * 1000
            
        Case PDFC_RED
            modifiedTolerance = m_Tolerance
            
        Case PDFC_GREEN
            modifiedTolerance = m_Tolerance
            
        Case PDFC_BLUE
            modifiedTolerance = m_Tolerance
            
        Case PDFC_ALPHA
            modifiedTolerance = m_Tolerance
    
    End Select
    
    Const FILLED_PIXEL As Long = &HFFFFFFFF
    
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
        quickX = x * 4
        b = srcImageData(quickX, y)
        g = srcImageData(quickX + 1, y)
        r = srcImageData(quickX + 2, y)
        a = srcImageData(quickX + 3, y)
    
        'Compare this pixel against the reference
        If (m_CompareMode = PDFC_COMPOSITE) Then
            thisValue = Abs(r - refR) + Abs(g - refG) + Abs(b - refB) + Abs(a - refA)
            isWithinTolerance = (thisValue <= modifiedTolerance)
        
        ElseIf (m_CompareMode = PDFC_COLOR) Then
            thisValue = Abs(r - refR) + Abs(g - refG) + Abs(b - refB)
            isWithinTolerance = (thisValue <= modifiedTolerance)
            
        ElseIf (m_CompareMode = PDFC_LUMINANCE) Then
            l = 213 * r + 715 * g + 72 * b
            isWithinTolerance = (Abs(l - refL) <= modifiedTolerance)
            
        ElseIf (m_CompareMode = PDFC_RED) Then
            isWithinTolerance = (Abs(r - refR) <= modifiedTolerance)
            
        ElseIf (m_CompareMode = PDFC_GREEN) Then
            isWithinTolerance = (Abs(g - refG) <= modifiedTolerance)
                
        ElseIf (m_CompareMode = PDFC_BLUE) Then
            isWithinTolerance = (Abs(b - refB) <= modifiedTolerance)
                
        ElseIf (m_CompareMode = PDFC_ALPHA) Then
            isWithinTolerance = (Abs(a - refA) <= modifiedTolerance)
        
        End If
        
        'If this value is within the requested tolerance, mark it on the destination map
        If isWithinTolerance Then
            
            dstImageData(x, y) = FILLED_PIXEL
            
            'If antialiasing is active, mark this pixel as filled; this simplifies the AA process
            If m_AntialiasingMode Then m_AlreadyChecked(x, y) = 2
        
            'Push our four neighboring pixels onto the stack, as necessary
            If (x > 0) Then If (m_AlreadyChecked(x - 1, y) = 0) Then PushOntoStack x - 1, y
            If (x < xBound) Then If (m_AlreadyChecked(x + 1, y) = 0) Then PushOntoStack x + 1, y
            
            If (y > 0) Then If (m_AlreadyChecked(x, y - 1) = 0) Then PushOntoStack x, y - 1
            If (y < yBound) Then If (m_AlreadyChecked(x, y + 1) = 0) Then PushOntoStack x, y + 1
            
        End If
                
    'As long as there are more stack points to process, rinse and repeat
    Loop While m_StackPosition >= 0
    
    'Release our array references
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    dstDIB.UnwrapLongArrayFromDIB dstImageData
    
    FloodFillContiguous = True
    
End Function

'Perform a full-image, non-contiguous flood fill.  Note that dstDIB is guaranteed to be properly initialized.
'
'IMPORTANT NOTE!  As of v7.0, both source and destination DIBs are required to be 32-bpp.  Passing a 24-bpp image will
' cause a hard crash.  (This matches PD's internal conversion to always-enforced 32-bpp sources.)
Private Function FloodFillGlobal(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Boolean
    
     'Predetermine upper bounds for x/y checks
    Dim xBound As Long, yBound As Long
    xBound = srcDIB.GetDIBWidth - 1
    yBound = srcDIB.GetDIBHeight - 1
    
    'Make sure 24 and 32bpp sources are both handled correctly
    Dim x As Long, y As Long, quickX As Long, dstQuickX As Long
    
    'Generate direct references to the source and destination DIB data
    Dim srcImageData() As Byte, dstImageData() As Long
    Dim srcSA As SAFEARRAY2D, dstSA As SAFEARRAY2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'For performance reasons, the destination DIB uses a Long-type array instead of a byte array (as we don't need access to individual bytes)
    dstDIB.WrapLongArrayAroundDIB dstImageData, dstSA
    
    'A number of local variables are used to help optimize the flood function
    Dim isWithinTolerance As Boolean
    Dim modifiedTolerance As Double
    
    'Populate our reference comparison values
    Dim r As Long, g As Long, b As Long, a As Long, l As Long
    Dim refR As Long, refG As Long, refB As Long, refA As Long, refL As Long
    Dim thisValue As Double
    
    quickX = m_InitPoint.x * 4
    y = m_InitPoint.y
    
    refB = srcImageData(quickX, y)
    refG = srcImageData(quickX + 1, y)
    refR = srcImageData(quickX + 2, y)
    refA = srcImageData(quickX + 3, y)
    
    refL = (213 * refR + 715 * refG + 72 * refB)
    
    'Calculate a reference tolerance value, which serves as the base for the flood fill
    If (m_CompareMode = PDFC_COMPOSITE) Then
        'Composite results do not require a base value, as they are independently processed against the reference
        ' RGBA values as we go.  However, to accelerate the required check, we premultiply the requested tolerance
        ' by 4, to avoid the need for a divide function in the inner loop
        modifiedTolerance = m_Tolerance * 4
            
    ElseIf (m_CompareMode = PDFC_COLOR) Then
        modifiedTolerance = m_Tolerance * 3
        
    ElseIf (m_CompareMode = PDFC_LUMINANCE) Then
        'To save time on the inner loop, we don't divide luminance by 1000; to make this work, we must change the
        ' tolerance range to [0, 1000] instead of [0, 255.0]
        modifiedTolerance = m_Tolerance * 1000
            
    ElseIf (m_CompareMode = PDFC_RED) Then
        modifiedTolerance = m_Tolerance
            
    ElseIf (m_CompareMode = PDFC_GREEN) Then
        modifiedTolerance = m_Tolerance
            
    ElseIf (m_CompareMode = PDFC_BLUE) Then
        modifiedTolerance = m_Tolerance
            
    ElseIf (m_CompareMode = PDFC_ALPHA) Then
        modifiedTolerance = m_Tolerance
    
    End If
    
    Const FILLED_PIXEL As Long = &HFFFFFFFF
    
    'Start processing the image!
    For x = 0 To xBound
        quickX = x * 4
    For y = 0 To yBound
    
        'Reset the tolerance check
        isWithinTolerance = False
        
        'Retrieve RGB/A values for this point
        b = srcImageData(quickX, y)
        g = srcImageData(quickX + 1, y)
        r = srcImageData(quickX + 2, y)
        a = srcImageData(quickX + 3, y)
    
        'Compare this pixel against the reference
        If (m_CompareMode = PDFC_COMPOSITE) Then
                thisValue = Abs(r - refR) + Abs(g - refG) + Abs(b - refB) + Abs(a - refA)
                isWithinTolerance = (thisValue <= modifiedTolerance)
            
        ElseIf (m_CompareMode = PDFC_COLOR) Then
                thisValue = Abs(r - refR) + Abs(g - refG) + Abs(b - refB)
                isWithinTolerance = (thisValue <= modifiedTolerance)
            
        ElseIf (m_CompareMode = PDFC_LUMINANCE) Then
                l = 213 * r + 715 * g + 72 * b
                isWithinTolerance = (Abs(l - refL) <= modifiedTolerance)
            
        ElseIf (m_CompareMode = PDFC_RED) Then
                isWithinTolerance = (Abs(r - refR) <= modifiedTolerance)
            
        ElseIf (m_CompareMode = PDFC_GREEN) Then
                isWithinTolerance = (Abs(g - refG) <= modifiedTolerance)
                
        ElseIf (m_CompareMode = PDFC_BLUE) Then
                isWithinTolerance = (Abs(b - refB) <= modifiedTolerance)
                
        ElseIf (m_CompareMode = PDFC_ALPHA) Then
                isWithinTolerance = (Abs(a - refA) <= modifiedTolerance)
        
        End If
        
        'If this value is within the requested tolerance, mark it on the destination map
        If isWithinTolerance Then dstImageData(x, y) = FILLED_PIXEL
        
    Next y
    Next x
        
    'Release our array references
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    dstDIB.UnwrapLongArrayFromDIB dstImageData
    
    FloodFillGlobal = True
    
End Function

'Stack helper functions
Private Sub PushOntoStack(ByVal x As Long, ByVal y As Long)
    
    m_StackPosition = m_StackPosition + 1
    
    'Resize the stack as necessary
    If (m_StackPosition > m_StackHeight) Then
        m_StackHeight = m_StackHeight * 2 + 1
        ReDim Preserve m_Stack(0 To m_StackHeight) As POINTAPI
    End If
    
    'Mark this point as "due to be checked", so it does not get re-checked
    m_AlreadyChecked(x, y) = 1
    
    'Add the point to the stack
    m_Stack(m_StackPosition).x = x
    m_Stack(m_StackPosition).y = y
    
End Sub

Private Sub PopFromStack(ByRef x As Long, ByRef y As Long)
    x = m_Stack(m_StackPosition).x
    y = m_Stack(m_StackPosition).y
    m_StackPosition = m_StackPosition - 1
End Sub

'Apply a custom antialiasing solution to a DIB.  The DIB must contain the contents of a *contiguous* flood fill operation.
' A global flood fill won't work, as it doesn't build a fill table.
Private Function FloodFillCustomAA(ByRef floodFillDIB As pdDIB) As Boolean
    
    'Predetermine upper bounds for x/y checks
    Dim xBound As Long, yBound As Long
    xBound = floodFillDIB.GetDIBWidth - 1
    yBound = floodFillDIB.GetDIBHeight - 1
    
    'Make sure 24 and 32bpp sources are both handled correctly
    Dim x As Long, y As Long, quickX As Long
    
    'Build a lookup table of already premultiplied values
    Dim fLookup() As Long
    ReDim fLookup(0 To 255) As Long
    
    Dim tmpHold As RGBQUAD
    
    For x = 0 To 255
        With tmpHold
            .alpha = x
            .Blue = x
            .Green = x
            .Red = x
        End With
        CopyMemory ByVal VarPtr(fLookup(x)), ByVal VarPtr(tmpHold), 4&
    Next x
    
    'Wrap a long-type array around the flood fill data
    Dim fImageData() As Long
    Dim fSA As SAFEARRAY2D
    floodFillDIB.WrapLongArrayAroundDIB fImageData, fSA
        
    'Populate our reference comparison values
    Dim aaStrength As Long, finalResult As Byte
    Dim yDownSafe As Boolean, yUpSafe As Boolean
    
    Dim constMultiplier As Double
    constMultiplier = (255 / 18) * 0.8
    
    'Start processing the image!
    For x = 0 To xBound
    For y = 0 To yBound
    
        'The floodfill map is used to calculate antialiasing.  If it contains a 1 at this location, it means this pixel
        ' was checked for the floodfill operation, but not actually filled.  That makes it a target for antialiasing.
        If (m_AlreadyChecked(x, y) = 1) Then
        
            aaStrength = 0
            
            'To simplify boundary checks, mark y safety in advance
            yDownSafe = (y > 0)
            yUpSafe = (y < yBound)
            
            'Check the eight surrounding pixels in the AA table.  If one was filled, its value will be 2.  If one was
            ' checked but unfilled, its value will be 1.  If a pixel was never checked, its value will be zero.
            ' We calculate AA by adding a value of 2 for each filled pixel, and a value of 1 for each filled pixel,
            ' and dividing the total by 18.  This provides maximum AA coverage of 88.8%, for a pixel surrounded on all
            ' eight sides by filled pixels.  (Note that after some testing, I decided to always reduce the calculated
            ' AA value by 20%.  This provides a softer AA effect, and hopefully more aesthetically pleasing results.)
            
            'Check the three pixels to the left (conditionally)
            If (x > 0) Then
                If yDownSafe Then aaStrength = aaStrength + m_AlreadyChecked(x - 1, y - 1)
                aaStrength = aaStrength + m_AlreadyChecked(x - 1, y)
                If yUpSafe Then aaStrength = aaStrength + m_AlreadyChecked(x - 1, y + 1)
            End If
            
            'Check the pixels above and below
            If yDownSafe Then aaStrength = aaStrength + m_AlreadyChecked(x, y - 1)
            If yUpSafe Then aaStrength = aaStrength + m_AlreadyChecked(x, y + 1)
            
            'Check the pixels to the right (conditionally)
            If (x < xBound) Then
                If yDownSafe Then aaStrength = aaStrength + m_AlreadyChecked(x + 1, y - 1)
                aaStrength = aaStrength + m_AlreadyChecked(x + 1, y)
                If yUpSafe Then aaStrength = aaStrength + m_AlreadyChecked(x + 1, y + 1)
            End If
            
            'We know the floodfill map at this position is unfilled, so we don't even need to retrieve the value to
            ' calculate proper AA.
            finalResult = aaStrength * constMultiplier
            fImageData(x, y) = fLookup(finalResult)
            
        End If
        
    Next y
    Next x
        
    'Release our array references
    floodFillDIB.UnwrapLongArrayFromDIB fImageData
    
    FloodFillCustomAA = True

End Function

Private Sub Class_Initialize()

    'Reset all stack values
    m_StackPosition = 0
    m_StackHeight = INITIAL_STACK_HEIGHT - 1
    ReDim m_Stack(0 To m_StackHeight) As POINTAPI
    
    'Reset our check array
    ReDim m_AlreadyChecked(0, 0) As Byte
    
    'Composite is the default tolerance mode
    m_CompareMode = PDFC_COMPOSITE
    
End Sub
