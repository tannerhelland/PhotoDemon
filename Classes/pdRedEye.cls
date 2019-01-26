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
'Copyright 2015-2019 by Tanner Helland
'Created: 31/December/16
'Last updated: 06/January/16
'Last update: rework flood fill functionality into red-eye region detection
'
'See FormRedEye for details on interacting with this class.
'
'Like any array-based tool, this class will be slow inside the IDE.  Please compile before using.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'A stack is used to track pixels that need to be checked.  This stack will be reused over and over again as each
' new region is detected, but to save processing time we don't reallocate it unless absolutely necessary.
Private m_Stack() As PointAPI
Private m_StackPosition As Long
Private m_StackHeight As Long
Private Const INITIAL_STACK_HEIGHT As Long = 4096

'Red-eye detection requires two input arrays.  In the first array, we will search already identified red-eye,
' eye-highlight, and non-skin pixels to determine potentially valid eye regions.  This array must have already
' been filled by the user.
Private m_HighlightData() As Byte, m_HighlightSA As SafeArray2D

'...The second input array is of INTEGER type.  It must be dimensioned to the size of the target image, and it will
' store unique region IDs for each pixel.  Because there may be more than 255 unique regions, a byte array won't work.
' Pixels that do not belong to any regions are marked as 0.
Private m_RegionIDs() As Integer, m_RegionSA As SafeArray2D

'Width and height of the target image.
Private m_Width As Long, m_Height As Long

'As each region is identified, we're going to add its data to a region stack.  This stack is completely different from
' the pixel coordinate stack used for region detection.  It stores additional information about each region,
' (like boundary positions) which greatly simplifies the subsequent steps of processing each region in turn.
Private m_RegionStack() As PD_Dynamic_Region
Private m_RegionStackPosition As Long
Private m_RegionStackHeight As Long
Private Const INITIAL_REGION_STACK_HEIGHT As Long = 256

'The current region ID.  The first region is marked as "1", and each subsequent region is assigned in ascending order.
Private m_CurrentRegion As Long

'Initiate red-eye detection.  This step simply aliases the target arrays supplied by the user, one each for region
' tracking and region detection, while also prepping some internal structs.
Friend Function InitializeRedEyeEngine(ByVal srcWidth As Long, ByVal srcHeight As Long, ByRef srcHighlights() As Byte, ByRef srcRegions() As Integer) As Boolean
    
    m_Width = srcWidth
    m_Height = srcHeight
    
    'Alias our internal arrays around the ones provided by the caller.  As you might expect, you MUST NOT attempt
    ' to erase or ReDim the target arrays until the red-eye detector has been shut down.
    VBHacks.Alias2DArray_Byte srcHighlights, m_HighlightData, m_HighlightSA
    VBHacks.Alias2DArray_Integer srcRegions, m_RegionIDs, m_RegionSA
    
    'Reset all internal tracking arrays
    m_RegionStackHeight = INITIAL_REGION_STACK_HEIGHT - 1
    ReDim m_RegionStack(0 To m_RegionStackHeight) As PD_Dynamic_Region
    m_RegionStackPosition = 0
    
    m_StackHeight = INITIAL_STACK_HEIGHT - 1
    ReDim m_Stack(0 To m_StackHeight) As PointAPI
    m_StackPosition = 0
    
    m_CurrentRegion = 0
    
    InitializeRedEyeEngine = True

End Function

'When the red-eye process is finished (due to any condition - success, error, etc), the caller MUST call this
' function to release our aliases to their local red-eye tracking arrays.
Friend Function ReleaseRedEyeEngine(ByRef srcHighlights() As Byte, ByRef srcRegions() As Integer) As Boolean
    
    'The only termination task that really matters is releasing our array aliases; if we don't do this, memory will
    ' either leak, or VB will straight-up crash.
    VBHacks.Unalias2DArray_Byte srcHighlights, m_HighlightData
    VBHacks.Unalias2DArray_Integer srcRegions, m_RegionIDs
    
    ReleaseRedEyeEngine = True
    
End Function

'Generate a new highlight region using a given pixel as the seed.  The seed pixel must not already be part of a region;
' if it is, this function will fail.
'
'If successful, this function will return the ID of the assigned region.
Friend Function FindHighlightRegion(ByVal inputX As Long, ByVal inputY As Long, ByVal regionMarkerID As Long) As Integer

    'Increment the region counter.
    ' (TODO: if this exceeds the upper bound of an integer, exit immediately; this is an insanely unlikely scenario,
    ' but I want to do some tests on actual images to see if 32k+ highlight regions can be artificially induced -
    ' if they can, I'll convert the region-detection code to use Longs instead of Integers.)
    m_CurrentRegion = m_CurrentRegion + 1
    
    'Use the current region index to calculate an actual region ID.  (We don't want this to be on the range [-1, 1]
    ' as those values have special meaning to the iterator.)
    Dim regID As Long
    regID = m_CurrentRegion + 1
    
    'Reset the stack.  Note that we don't actually resize the stack; this is an optimization technique to improve performance
    ' if this class is used multiple times in a row.
    m_StackPosition = -1
    
    'Predetermine upper bounds for x/y checks
    Dim xBound As Long, yBound As Long
    xBound = m_Width - 1
    yBound = m_Height - 1
    
    'For each region, we also track the region's boundaries.  (This is easiest to do here, rather than waste time
    ' retroactively calculating it in subsequent steps.)
    Dim regLeft As Long, regTop As Long, regRight As Long, regBottom As Long
    regLeft = xBound
    regTop = yBound
    regRight = 0
    regBottom = 0
    
    'We'll also count the number of pixels inside this region; this simplifies things like rejecting regions that
    ' are too large or too small.
    Dim regPixelCount As Long
    regPixelCount = 0
    
    'Populate the initial stack point
    Dim x As Long, y As Long
    PushOntoStack inputX, inputY
    
    'Start processing the stack
    Do

        'Retrieve the next point from the stack
        PopFromStack x, y
        
        'If this is a highlight pixel, add it to the current region ID
        If m_HighlightData(x, y) = regionMarkerID Then
            
            m_RegionIDs(x, y) = regID
            
            'Update boundaries
            If x < regLeft Then regLeft = x
            If x > regRight Then regRight = x
            If y < regTop Then regTop = y
            If y > regBottom Then regBottom = y
            
            'Update pixel count
            regPixelCount = regPixelCount + 1
            
            'If any of the four neighboring pixels have not yet been...
            ' 1) assigned to a region, and...
            ' 2) ...already added to the stack...
            ' ...add them to the stack now.
            If x > 0 Then If m_RegionIDs(x - 1, y) = 0 Then PushOntoStack x - 1, y
            If x < xBound Then If m_RegionIDs(x + 1, y) = 0 Then PushOntoStack x + 1, y
            If y > 0 Then If m_RegionIDs(x, y - 1) = 0 Then PushOntoStack x, y - 1
            If y < yBound Then If m_RegionIDs(x, y + 1) = 0 Then PushOntoStack x, y + 1
            
        End If
        
    'As long as there are more stack points to process, rinse and repeat
    Loop While m_StackPosition >= 0
    
    'With the stack cleared, we now know everything there is to know about this region.  Add the region's ID and
    ' boundaries to our stack collection.
    If m_RegionStackPosition > m_RegionStackHeight Then
        m_RegionStackHeight = m_RegionStackHeight * 2 + 1
        ReDim Preserve m_RegionStack(0 To m_RegionStackHeight) As PD_Dynamic_Region
    End If
    
    With m_RegionStack(m_RegionStackPosition)
        .RegionID = regID
        .RegionValid = True
        .RegionLeft = regLeft
        .RegionTop = regTop
        .RegionWidth = regRight - regLeft
        .RegionHeight = regBottom - regTop
        .RegionPixelCount = regPixelCount
        .SeedPixelX = inputX
        .SeedPixelY = inputY
    End With
    
    'Increment the region count and exit
    m_RegionStackPosition = m_RegionStackPosition + 1
    
    FindHighlightRegion = m_CurrentRegion
    
End Function

'Given an existing highlight region, expand it to include any neighboring red-eye pixels.  The end result is a region that
' encompasses both highlight and red-eye pixels.
Friend Function ExpandToIncludeRedEye(ByRef srcRegion As PD_Dynamic_Region, ByVal highlightID As Long, ByVal redEyeID As Long) As Boolean
    
    'Reset the stack.  Note that we don't actually resize the stack; this is an optimization technique to improve performance
    ' if this class is used multiple times in a row.
    m_StackPosition = -1
    
    'Predetermine upper bounds for x/y checks
    Dim xBound As Long, yBound As Long
    xBound = m_Width - 1
    yBound = m_Height - 1
    
    'Set the region dimension trackers to match the source region values
    Dim regLeft As Long, regTop As Long, regRight As Long, regBottom As Long
    regLeft = srcRegion.RegionLeft
    regTop = srcRegion.RegionTop
    regRight = srcRegion.RegionLeft + srcRegion.RegionWidth
    regBottom = srcRegion.RegionTop + srcRegion.RegionHeight
    
    'We're also going to generate a new pixel count for the region, one that accurately reflects the new
    ' highlight + red-eye criteria for pixel inclusion.
    Dim regPixelCount As Long
    regPixelCount = 0
    
    Dim regID As Long
    regID = srcRegion.RegionID
    
    'Populate the initial stack point
    Dim x As Long, y As Long
    PushOntoStack srcRegion.SeedPixelX, srcRegion.SeedPixelY
    
    'Start processing the stack
    Do
        
        'Retrieve the next point from the stack
        PopFromStack x, y
        
        'If this is a highlight pixel OR a red-eye pixel, assign it the current region ID
        If (m_HighlightData(x, y) = highlightID) Or (m_HighlightData(x, y) = redEyeID) Then
            
            m_RegionIDs(x, y) = regID
            
            'Update boundaries
            If x < regLeft Then regLeft = x
            If x > regRight Then regRight = x
            If y < regTop Then regTop = y
            If y > regBottom Then regBottom = y
            
            'Update pixel count
            regPixelCount = regPixelCount + 1
            
            'If any of the four neighboring pixels have not yet been...
            ' 1) assigned to a region, and...
            ' 2) ...already added to the stack...
            ' ...add them to the stack now.
            If x > 0 Then If m_RegionIDs(x - 1, y) = 0 Then PushOntoStack x - 1, y
            If x < xBound Then If m_RegionIDs(x + 1, y) = 0 Then PushOntoStack x + 1, y
            If y > 0 Then If m_RegionIDs(x, y - 1) = 0 Then PushOntoStack x, y - 1
            If y < yBound Then If m_RegionIDs(x, y + 1) = 0 Then PushOntoStack x, y + 1
            
        End If
        
    'As long as there are more stack points to process, rinse and repeat
    Loop While m_StackPosition >= 0
    
    'Once the region has been assembled, we need to update the source region object to match our new region data.
    With srcRegion
        .RegionLeft = regLeft
        .RegionTop = regTop
        .RegionWidth = regRight - regLeft
        .RegionHeight = regBottom - regTop
        .RegionPixelCount = regPixelCount
    End With
    
    ExpandToIncludeRedEye = True
    
End Function

'Return a copy of the current region stack.  If no regions were detected, this function will fail.
Friend Function GetCopyOfRegionStack(ByRef dstArray() As PD_Dynamic_Region, ByRef numOfRegions As Long) As Boolean
    
    numOfRegions = m_RegionStackPosition
    
    If numOfRegions > 0 Then
        ReDim dstArray(0 To numOfRegions - 1) As PD_Dynamic_Region
        CopyMemoryStrict VarPtr(dstArray(0)), VarPtr(m_RegionStack(0)), LenB(m_RegionStack(0)) * numOfRegions
        GetCopyOfRegionStack = True
    Else
        GetCopyOfRegionStack = False
    End If

End Function

'Stack helper functions
Private Sub PushOntoStack(ByVal x As Long, ByVal y As Long)
    
    m_StackPosition = m_StackPosition + 1

    'Resize the stack as necessary
    If m_StackPosition > m_StackHeight Then
        m_StackHeight = m_StackHeight * 2 + 1
        ReDim Preserve m_Stack(0 To m_StackHeight) As PointAPI
    End If
    
    'Mark this point as "due to be checked", so it does not get re-checked
    m_RegionIDs(x, y) = -1

    'Add the point to the stack
    With m_Stack(m_StackPosition)
        .x = x
        .y = y
    End With

End Sub

Private Sub PopFromStack(ByRef x As Long, ByRef y As Long)

    With m_Stack(m_StackPosition)
        x = .x
        y = .y
    End With

    m_StackPosition = m_StackPosition - 1

End Sub

'Given an RGB component and properly calibrated correction values, return a corrected value
Friend Function FixRedEyeColor(ByVal origValue As Long, ByVal brightCorrection As Double, ByVal contrastCorrection As Double, ByVal meanLight As Long) As Long
    
    'Apply brightness correction
    FixRedEyeColor = (0.4 * brightCorrection * 255 + origValue)
    
    'Clamp to byte range
    If FixRedEyeColor < 0 Then
        FixRedEyeColor = 0
    ElseIf FixRedEyeColor > 255 Then
        FixRedEyeColor = 255
    End If
    
    'Apply contrast correction on the brightened result
    FixRedEyeColor = FixRedEyeColor + (FixRedEyeColor - meanLight) * contrastCorrection
    
    'Clamp to byte range
    If FixRedEyeColor < 0 Then
        FixRedEyeColor = 0
    ElseIf FixRedEyeColor > 255 Then
        FixRedEyeColor = 255
    End If
    
End Function
