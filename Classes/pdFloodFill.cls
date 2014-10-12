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
'Copyright ©2013-2014 by Tanner Helland
'Created: 11/October/14
'Last updated: 11/October/14
'Last update: initial build
'
'Comments forthcoming...
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

'To reduce iterations, this function tracks pixels that have already been added to the stack
Private m_AlreadyChecked() As Byte

'Tolerance allows the user to control the strength of the flood
Private m_Tolerance As Double

'Different compare modes can be used to obtain better results.
Public Enum PD_FLOOD_COMPARE
    PDFC_COMPOSITE = 0
    PDFC_LUMINANCE = 1
    PDFC_RED = 2
    PDFC_GREEN = 3
    PDFC_BLUE = 4
    PDFC_ALPHA = 5
End Enum

#If False Then
    Const PDFC_COMPOSITE = 0, PDFC_LUMINANCE = 1, PDFC_RED = 2, PDFC_GREEN = 3, PDFC_BLUE = 4, PDFC_ALPHA = 5
#End If

Private m_CompareMode As PD_FLOOD_COMPARE

'Two search methods are supported: contiguous region based on the initial point (default behavior), or the full image
Public Enum PD_FLOOD_SEARCH
    PDFS_CONTIGUOUS = 0
    PDFS_WHOLEIMAGE = 1
End Enum

#If False Then
    Const PDFS_CONTIGUOUS = 0, PDFS_WHOLEIMAGE = 1
#End If

Private m_SearchMode As PD_FLOOD_SEARCH

'A custom antialiasing technique can be used to soften the floodfill results
Private m_AntialiasingMode As Boolean

'Obviously the function needs a starting x/y position
Private m_InitPoint As POINTAPI

'Get/set functions for all relevant flood fill parameters
Public Function getAntialiasingMode() As Boolean
    getAntialiasingMode = m_AntialiasingMode
End Function

Public Sub setAntialiasingMode(ByVal newAntialiasingMode As Boolean)
    m_AntialiasingMode = newAntialiasingMode
End Sub

Public Function getCompareMode() As PD_FLOOD_COMPARE
    getCompareMode = m_CompareMode
End Function

Public Function setCompareMode(ByVal newCompareMode As PD_FLOOD_COMPARE)
    m_CompareMode = newCompareMode
End Function

Friend Function getInitialPoint() As POINTAPI
    getInitialPoint = m_InitPoint
End Function

Public Sub setInitialPoint(ByVal startX As Long, ByVal startY As Long)
    m_InitPoint.x = startX
    m_InitPoint.y = startY
End Sub

Public Function getSearchMode() As PD_FLOOD_SEARCH
    getSearchMode = m_SearchMode
End Function

Public Sub setSearchMode(ByVal newSearchMode As PD_FLOOD_SEARCH)
    m_SearchMode = newSearchMode
End Sub

Public Function getTolerance() As Double
    getTolerance = m_Tolerance
End Function

Public Sub setTolerance(ByVal newTolerance As Double)
    m_Tolerance = newTolerance
End Sub

'Initiate a flood fill operation.  This class doesn't actually fill anything; what it does is fill a destination DIB with a
' graymap of the flood results (black = unfilled, white = filled, gray = partially filled).  This allows the caller to
' use the flood results however they want, without having to modify this class to match.
Public Function initiateFloodFill(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Boolean

    'Initialize the destination DIB
    dstDIB.createBlank srcDIB.getDIBWidth, srcDIB.getDIBHeight, 24, 0
    
    'Make sure the passed x/y coords are valid.  If they aren't, exit now.
    If (m_InitPoint.x < 0) Or (m_InitPoint.y < 0) Or (m_InitPoint.x >= srcDIB.getDIBWidth) Or (m_InitPoint.y >= srcDIB.getDIBHeight) Then
        Debug.Print "Invalid flood fill location requested.  Abandoning flood fill now."
        initiateFloodFill = False
        Exit Function
    End If
    
    'Based on the specified search mode, call the appropriate flood function
    Select Case m_SearchMode
        
        Case PDFS_CONTIGUOUS
            floodFillContiguous srcDIB, dstDIB
        
        Case PDFS_WHOLEIMAGE
            floodFillGlobal srcDIB, dstDIB
        
    End Select
    
    initiateFloodFill = True

End Function

'Perform a contiguous (default) flood fill.  Note that dstDIB is guaranteed to be properly initialized.
Private Function floodFillContiguous(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Boolean

    'Reset the stack.  Note that we don't actually resize the stack; this is an optimization technique to improve performance
    ' if this class is used multiple times in a row.
    m_StackPosition = -1
    
    'Predetermine upper bounds for x/y checks
    Dim xBound As Long, yBound As Long
    xBound = srcDIB.getDIBWidth - 1
    yBound = srcDIB.getDIBHeight - 1
    
    'Reset the "has been filled" array
    ReDim m_AlreadyChecked(0 To xBound, 0 To yBound) As Byte
    
    'Populate the initial stack point
    pushOntoStack m_InitPoint.x, m_InitPoint.y
    
    'Make sure 24 and 32bpp sources are both handled correctly
    Dim x As Long, y As Long, QuickX As Long, xOffset As Long, hasAlpha As Boolean
    xOffset = srcDIB.getDIBColorDepth \ 8
    If srcDIB.getDIBColorDepth = 32 Then hasAlpha = True Else hasAlpha = False
    
    'Generate direct references to the source and destination DIB data
    Dim srcImageData() As Byte, dstImageData() As Byte
    Dim srcSA As SAFEARRAY2D, dstSA As SAFEARRAY2D
    
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'A number of local variables are used to help optimize the flood function
    Dim tmpPoint As POINTAPI
    Dim isWithinTolerance As Boolean
    Dim modifiedTolerance As Double
    
    'Populate our reference comparison values
    Dim r As Long, g As Long, b As Long, a As Long, l As Long
    Dim refR As Long, refG As Long, refB As Long, refA As Long, refL As Long
    
    Dim referenceValue As Double, thisValue As Double
    
    QuickX = m_InitPoint.x * xOffset
    y = m_InitPoint.y
    
    refR = srcImageData(QuickX + 2, y)
    refG = srcImageData(QuickX + 1, y)
    refB = srcImageData(QuickX, y)
    If hasAlpha Then refA = srcImageData(QuickX + 3, y) Else refA = 0
    
    refL = 213 * refR + 715 * refG + 72 * refB
    
    'Calculate a reference tolerance value, which serves as the base for the flood fill
    Select Case m_CompareMode
    
        Case PDFC_COMPOSITE
            
            'Composite results do not require a base value, as they are independently processed against the reference
            ' RGB values as we go.  However, to accelerate the required check, we premultiply the requested tolerance
            ' by 3, to avoid the need for a divide function in the inner loop
            If hasAlpha Then
                modifiedTolerance = m_Tolerance * 4
            Else
                modifiedTolerance = m_Tolerance * 3
            End If
        
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
    
    'Start processing the stack!
    Do
    
        'Reset the tolerance check
        isWithinTolerance = False
        
        'Retrieve the next point from the stack
        popFromStack x, y
                
        'Retrieve RGB/A values for this point
        QuickX = x * xOffset
        r = srcImageData(QuickX + 2, y)
        g = srcImageData(QuickX + 1, y)
        b = srcImageData(QuickX, y)
        If hasAlpha Then a = srcImageData(QuickX + 3, y)
    
        'Compare this pixel against the reference
        Select Case m_CompareMode
                
            Case PDFC_COMPOSITE
                thisValue = Abs(r - refR) + Abs(g - refG) + Abs(b - refB)
                If hasAlpha Then thisValue = thisValue + Abs(a - refA)
                If thisValue <= modifiedTolerance Then isWithinTolerance = True
            
            Case PDFC_LUMINANCE
                l = 213 * r + 715 * g + 72 * b
                If Abs(l - refL) <= modifiedTolerance Then isWithinTolerance = True
            
            Case PDFC_RED
                If Abs(r - refR) <= modifiedTolerance Then isWithinTolerance = True
            
            Case PDFC_GREEN
                If Abs(g - refG) <= modifiedTolerance Then isWithinTolerance = True
                
            Case PDFC_BLUE
                If Abs(b - refB) <= modifiedTolerance Then isWithinTolerance = True
                
            Case PDFC_ALPHA
                If Abs(a - refA) <= modifiedTolerance Then isWithinTolerance = True
        
        End Select
        
        'If this value is within the requested tolerance, mark it on the destination map
        If isWithinTolerance Then
        
            dstImageData(x * 3, y) = 255
            
            'If antialiasing is active, mark this pixel as filled; this simplifies the AA process
            If m_AntialiasingMode Then m_AlreadyChecked(x, y) = 2
        
            'Push our four neighboring pixels onto the stack, as necessary
            If x > 0 Then If m_AlreadyChecked(x - 1, y) = 0 Then pushOntoStack x - 1, y
            If x < xBound Then If m_AlreadyChecked(x + 1, y) = 0 Then pushOntoStack x + 1, y
            
            If y > 0 Then If m_AlreadyChecked(x, y - 1) = 0 Then pushOntoStack x, y - 1
            If y < yBound Then If m_AlreadyChecked(x, y + 1) = 0 Then pushOntoStack x, y + 1
            
        End If
                
    'As long as there are more stack points to process, rinse and repeat
    Loop While m_StackPosition >= 0
    
    'Release our array references
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    floodFillContiguous = True
    
End Function

'Perform a full-image, non-contiguous flood fill.  Note that dstDIB is guaranteed to be properly initialized.
Private Function floodFillGlobal(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Boolean
    
     'Predetermine upper bounds for x/y checks
    Dim xBound As Long, yBound As Long
    xBound = srcDIB.getDIBWidth - 1
    yBound = srcDIB.getDIBHeight - 1
    
    'Make sure 24 and 32bpp sources are both handled correctly
    Dim x As Long, y As Long, QuickX As Long, xOffset As Long, hasAlpha As Boolean
    xOffset = srcDIB.getDIBColorDepth \ 8
    If srcDIB.getDIBColorDepth = 32 Then hasAlpha = True Else hasAlpha = False
    
    'Generate direct references to the source and destination DIB data
    Dim srcImageData() As Byte, dstImageData() As Byte
    Dim srcSA As SAFEARRAY2D, dstSA As SAFEARRAY2D
    
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'A number of local variables are used to help optimize the flood function
    Dim isWithinTolerance As Boolean
    Dim modifiedTolerance As Double
    
    'Populate our reference comparison values
    Dim r As Long, g As Long, b As Long, a As Long, l As Long
    Dim refR As Long, refG As Long, refB As Long, refA As Long, refL As Long
    
    Dim referenceValue As Double, thisValue As Double
    
    QuickX = m_InitPoint.x * xOffset
    y = m_InitPoint.y
    
    refR = srcImageData(QuickX + 2, y)
    refG = srcImageData(QuickX + 1, y)
    refB = srcImageData(QuickX, y)
    If hasAlpha Then refA = srcImageData(QuickX + 3, y) Else refA = 0
    
    refL = (213 * refR + 715 * refG + 72 * refB)
    
    'Calculate a reference tolerance value, which serves as the base for the flood fill
    Select Case m_CompareMode
    
        Case PDFC_COMPOSITE
            
            'Composite results do not require a base value, as they are independently processed against the reference
            ' RGB values as we go.  However, to accelerate the required check, we premultiply the requested tolerance
            ' by 3, to avoid the need for a divide function in the inner loop
            If hasAlpha Then
                modifiedTolerance = m_Tolerance * 4
            Else
                modifiedTolerance = m_Tolerance * 3
            End If
        
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
    
    'Start processing the image!
    For x = 0 To xBound
        QuickX = x * xOffset
    For y = 0 To yBound
    
        'Reset the tolerance check
        isWithinTolerance = False
        
        'Retrieve RGB/A values for this point
        r = srcImageData(QuickX + 2, y)
        g = srcImageData(QuickX + 1, y)
        b = srcImageData(QuickX, y)
        If hasAlpha Then a = srcImageData(QuickX + 3, y)
    
        'Compare this pixel against the reference
        Select Case m_CompareMode
                
            Case PDFC_COMPOSITE
                thisValue = Abs(r - refR) + Abs(g - refG) + Abs(b - refB)
                If hasAlpha Then thisValue = thisValue + Abs(a - refA)
                If thisValue <= modifiedTolerance Then isWithinTolerance = True
            
            Case PDFC_LUMINANCE
                l = 213 * r + 715 * g + 72 * b
                If Abs(l - refL) <= modifiedTolerance Then isWithinTolerance = True
            
            Case PDFC_RED
                If Abs(r - refR) <= modifiedTolerance Then isWithinTolerance = True
            
            Case PDFC_GREEN
                If Abs(g - refG) <= modifiedTolerance Then isWithinTolerance = True
                
            Case PDFC_BLUE
                If Abs(b - refB) <= modifiedTolerance Then isWithinTolerance = True
                
            Case PDFC_ALPHA
                If Abs(a - refA) <= modifiedTolerance Then isWithinTolerance = True
        
        End Select
        
        'If this value is within the requested tolerance, mark it on the destination map
        If isWithinTolerance Then dstImageData(x * 3, y) = 255
        
    Next y
    Next x
        
    'Release our array references
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    floodFillGlobal = True
    
End Function

'Stack helper functions
Private Sub pushOntoStack(ByVal x As Long, ByVal y As Long)
    
    m_StackPosition = m_StackPosition + 1
    
    'Resize the stack as necessary
    If m_StackPosition > m_StackHeight Then
        m_StackHeight = m_StackHeight * 2 + 1
        ReDim Preserve m_Stack(0 To m_StackHeight) As POINTAPI
    End If
    
    'Mark this point as "due to be checked", so it does not get re-checked
    m_AlreadyChecked(x, y) = 1
    
    'Add the point to the stack
    With m_Stack(m_StackPosition)
        .x = x
        .y = y
    End With

End Sub

Private Sub popFromStack(ByRef x As Long, ByRef y As Long)
    
    With m_Stack(m_StackPosition)
        x = .x
        y = .y
    End With
    
    m_StackPosition = m_StackPosition - 1
    
End Sub

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
