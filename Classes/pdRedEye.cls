VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdRedEye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Red-eye Detection Engine
'Copyright 2015-2016 by Tanner Helland
'Created: 31/December/16
'Last updated: 06/January/16
'Last update: rework flood fill functionality into red-eye region detection
'
'See FormRedEye for details on interacting with this class.
'
'Like any array-based tool, this class will be slow inside the IDE.  Please use only when compiled if at all possible.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'A stack is used to track pixels that need to be checked.  This stack will be reused over and over again as each
' new region is detected, but to save processing time we don't reallocate it unless absolutely necessary.
Private m_Stack() As POINTAPI
Private m_StackPosition As Long
Private m_StackHeight As Long
Private Const INITIAL_STACK_HEIGHT As Long = 4096

'Red-eye detection requires two input arrays.  In the first array, we will search already identified red-eye,
' eye-highlight, and non-skin pixels to determine potentially valid eye regions.  This array must have already
' been filled by the user.
Private m_HighlightData() As Byte, m_HighlightSA As SAFEARRAY2D

'...The second input array is of INTEGER type.  It must be dimensioned to the size of the target image, and it will
' store unique region IDs for each pixel.  Because there may be more than 255 unique regions, a byte array won't work.
' Pixels that do not belong to any regions are marked as 0.
Private m_RegionIDs() As Integer, m_RegionSA As SAFEARRAY2D

'Width and height of the target image.
Private m_Width As Long, m_Height As Long

'As each region is identified, we're going to add its data to a region stack.  This stack is completely different from
' the pixel coordinate stack used for region detection.  It stores the boundaries of the current region, which greatly
' simplifies the subsequent step of processing each region in turn.
Private Type PD_DYNAMIC_REGION
    RegionID As Long
    RegionLeft As Long
    RegionTop As Long
    RegionWidth As Long
    RegionHeight As Long
End Type

Private m_RegionStack() As PD_DYNAMIC_REGION
Private m_RegionStackPosition As Long
Private m_RegionStackHeight As Long
Private Const INITIAL_REGION_STACK_HEIGHT As Long = 256

'Obviously the function needs a starting x/y position
Private m_InitPoint As POINTAPI

'Initiate red-eye detection.  This step simply aliases the target arrays supplied by the user, one each for region
' tracking and region detection, while also prepping some internal structs.
Public Function InitializeRedEyeEngine(ByVal srcWidth As Long, ByVal srcHeight As Long, ByRef srcHighlights() As Byte, ByRef srcRegions() As Integer) As Boolean
    
    m_Width = srcWidth
    m_Height = srcHeight
    
    'Alias our internal arrays around the ones provided by the caller.  As you might expect, you MUST NOT attempt
    ' to erase or ReDim the target arrays until the red-eye detector has been shut down.
    VB_Hacks.Alias2DArray_Byte srcHighlights, m_HighlightData, m_HighlightSA
    VB_Hacks.Alias2DArray_Integer srcRegions, m_RegionIDs, m_RegionSA
    
    'Reset all internal tracking arrays
    m_RegionStackHeight = INITIAL_REGION_STACK_HEIGHT - 1
    ReDim m_RegionStack(0 To m_RegionStackHeight) As PD_DYNAMIC_REGION
    m_RegionStackPosition = 0
    
    m_StackHeight = INITIAL_STACK_HEIGHT - 1
    ReDim m_Stack(0 To m_StackHeight) As POINTAPI
    m_StackPosition = 0
    
    InitializeRedEyeEngine = True

End Function

'TEMPORARY CODE ONLY: we're going to borrow PD's heavily optimized floodfill function as the basis for our region
' detection code.  I'm leaving this for a single commit as a reminder of where this code originated.
'
''Perform a contiguous (default) flood fill.  Note that dstDIB is guaranteed to be properly initialized.
'Private Function floodFillContiguous(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Boolean
'
'    'Reset the stack.  Note that we don't actually resize the stack; this is an optimization technique to improve performance
'    ' if this class is used multiple times in a row.
'    m_StackPosition = -1
'
'    'Predetermine upper bounds for x/y checks
'    Dim xBound As Long, yBound As Long
'    xBound = srcDIB.getDIBWidth - 1
'    yBound = srcDIB.getDIBHeight - 1
'
'    'Reset the "has been filled" array
'    ReDim m_AlreadyChecked(0 To xBound, 0 To yBound) As Byte
'
'    'Populate the initial stack point
'    PushOntoStack m_InitPoint.x, m_InitPoint.y
'
'    'Make sure 24 and 32bpp sources are both handled correctly
'    Dim x As Long, y As Long, quickX As Long, dstQuickX As Long, xOffset As Long, hasAlpha As Boolean
'    xOffset = srcDIB.getDIBColorDepth \ 8
'    If srcDIB.getDIBColorDepth = 32 Then hasAlpha = True Else hasAlpha = False
'
'    'Generate direct references to the source and destination DIB data
'    Dim srcImageData() As Byte, dstImageData() As Byte
'    Dim srcSA As SAFEARRAY2D, dstSA As SAFEARRAY2D
'
'    prepSafeArray srcSA, srcDIB
'    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
'
'    prepSafeArray dstSA, dstDIB
'    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
'
'    'A number of local variables are used to help optimize the flood function
'    Dim tmpPoint As POINTAPI
'    Dim isWithinTolerance As Boolean
'    Dim modifiedTolerance As Double
'
'    'Populate our reference comparison values
'    Dim r As Long, g As Long, b As Long, a As Long, l As Long
'    Dim refR As Long, refG As Long, refB As Long, refA As Long, refL As Long
'
'    Dim referenceValue As Double, thisValue As Double
'
'    quickX = m_InitPoint.x * xOffset
'    y = m_InitPoint.y
'
'    refR = srcImageData(quickX + 2, y)
'    refG = srcImageData(quickX + 1, y)
'    refB = srcImageData(quickX, y)
'    If hasAlpha Then refA = srcImageData(quickX + 3, y) Else refA = 0
'
'    refL = 213 * refR + 715 * refG + 72 * refB
'
'    'Calculate a reference tolerance value, which serves as the base for the flood fill
'    Select Case m_CompareMode
'
'        Case PDFC_COMPOSITE
'
'            'Composite results do not require a base value, as they are independently processed against the reference
'            ' RGB values as we go.  However, to accelerate the required check, we premultiply the requested tolerance
'            ' by 3, to avoid the need for a divide function in the inner loop
'            If hasAlpha Then
'                modifiedTolerance = m_Tolerance * 4
'            Else
'                modifiedTolerance = m_Tolerance * 3
'            End If
'
'        Case PDFC_COLOR
'            modifiedTolerance = m_Tolerance * 3
'
'        Case PDFC_LUMINANCE
'
'            'To save time on the inner loop, we don't divide luminance by 1000; to make this work, we must change the
'            ' tolerance range to [0, 1000] instead of [0, 255.0]
'            modifiedTolerance = m_Tolerance * 1000
'
'        Case PDFC_RED
'            modifiedTolerance = m_Tolerance
'
'        Case PDFC_GREEN
'            modifiedTolerance = m_Tolerance
'
'        Case PDFC_BLUE
'            modifiedTolerance = m_Tolerance
'
'        Case PDFC_ALPHA
'            modifiedTolerance = m_Tolerance
'
'    End Select
'
'    'Start processing the stack!
'    Do
'
'        'Reset the tolerance check
'        isWithinTolerance = False
'
'        'Retrieve the next point from the stack
'        PopFromStack x, y
'
'        'Retrieve RGB/A values for this point
'        quickX = x * xOffset
'        r = srcImageData(quickX + 2, y)
'        g = srcImageData(quickX + 1, y)
'        b = srcImageData(quickX, y)
'        If hasAlpha Then a = srcImageData(quickX + 3, y)
'
'        'Compare this pixel against the reference
'        Select Case m_CompareMode
'
'            Case PDFC_COMPOSITE
'                thisValue = Abs(r - refR) + Abs(g - refG) + Abs(b - refB)
'                If hasAlpha Then thisValue = thisValue + Abs(a - refA)
'                If thisValue <= modifiedTolerance Then isWithinTolerance = True
'
'            Case PDFC_COLOR
'                thisValue = Abs(r - refR) + Abs(g - refG) + Abs(b - refB)
'                If thisValue <= modifiedTolerance Then isWithinTolerance = True
'
'            Case PDFC_LUMINANCE
'                l = 213 * r + 715 * g + 72 * b
'                If Abs(l - refL) <= modifiedTolerance Then isWithinTolerance = True
'
'            Case PDFC_RED
'                If Abs(r - refR) <= modifiedTolerance Then isWithinTolerance = True
'
'            Case PDFC_GREEN
'                If Abs(g - refG) <= modifiedTolerance Then isWithinTolerance = True
'
'            Case PDFC_BLUE
'                If Abs(b - refB) <= modifiedTolerance Then isWithinTolerance = True
'
'            Case PDFC_ALPHA
'                If Abs(a - refA) <= modifiedTolerance Then isWithinTolerance = True
'
'        End Select
'
'        'If this value is within the requested tolerance, mark it on the destination map
'        If isWithinTolerance Then
'
'            dstQuickX = x * 3
'            dstImageData(dstQuickX, y) = 255
'            dstImageData(dstQuickX + 1, y) = 255
'            dstImageData(dstQuickX + 2, y) = 255
'
'            'If antialiasing is active, mark this pixel as filled; this simplifies the AA process
'            If m_AntialiasingMode Then m_AlreadyChecked(x, y) = 2
'
'            'Push our four neighboring pixels onto the stack, as necessary
'            If x > 0 Then If m_AlreadyChecked(x - 1, y) = 0 Then PushOntoStack x - 1, y
'            If x < xBound Then If m_AlreadyChecked(x + 1, y) = 0 Then PushOntoStack x + 1, y
'
'            If y > 0 Then If m_AlreadyChecked(x, y - 1) = 0 Then PushOntoStack x, y - 1
'            If y < yBound Then If m_AlreadyChecked(x, y + 1) = 0 Then PushOntoStack x, y + 1
'
'        End If
'
'    'As long as there are more stack points to process, rinse and repeat
'    Loop While m_StackPosition >= 0
'
'    'Release our array references
'    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
'    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
'
'    floodFillContiguous = True
'
'End Function

''Stack helper functions
'Private Sub PushOntoStack(ByVal x As Long, ByVal y As Long)
'
'    m_StackPosition = m_StackPosition + 1
'
'    'Resize the stack as necessary
'    If m_StackPosition > m_StackHeight Then
'        m_StackHeight = m_StackHeight * 2 + 1
'        ReDim Preserve m_Stack(0 To m_StackHeight) As POINTAPI
'    End If
'
'    'Mark this point as "due to be checked", so it does not get re-checked
'    m_AlreadyChecked(x, y) = 1
'
'    'Add the point to the stack
'    With m_Stack(m_StackPosition)
'        .x = x
'        .y = y
'    End With
'
'End Sub
'
'Private Sub PopFromStack(ByRef x As Long, ByRef y As Long)
'
'    With m_Stack(m_StackPosition)
'        x = .x
'        y = .y
'    End With
'
'    m_StackPosition = m_StackPosition - 1
'
'End Sub
