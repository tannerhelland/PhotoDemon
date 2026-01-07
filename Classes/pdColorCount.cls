VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdColorCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Fast unique color counting class
'Copyright 2019-2026 by Tanner Helland
'Created: 26/March/19
'Last updated: 01/April/19
'Last update: optimize memory usage by treating alpha = 255 as a special case; this avoids allocating
'             leaf nodes for colors that only occur with full opacity, which reduces memory requirements
'             significantly regardless of whether the caller requests RGB vs RGBA tracking.
'
'This class provides a fast, low-resource way to count unique colors in an image.  Unique RGB triplets
' and RGBA quads are both counted "for free", although RGBA tracking can be forcibly disabled if you
' know an image is opaque.  (This makes little difference to resource usage or performance, and is
' provided more as a convenience).
'
'All discovered colors can be very quickly returned as a list of RGBQuads, regardless of final color
' count or tree size.
'
'The class is implemented as a specialized b-tree, with a standard VB array underlying the actual
' data storage.  A pure class-based implementation was tested at first, but construction/destruction
' times for the occasional 1+ million classes was a nightmare.  This new array+custom type
' implementation is multiple orders of magnitude faster (and significantly lighter on resources).
'
'Memory usage is quite good, even on images with large (1+ mln) counts of unique RGBA quads.
' I've even tested it on a specially constructed PNG image with all 16.7 million possible RGB triplets
' (4096x4096 pixels), and memory usage of the completed tree comes out to ~72.3 mb.  Manually adding a
' row of transparent pixels to the image (to force full RGBA tracking mode) increased memory usage to
' ~72.8 mb.  That's pretty awesome considering that the tree stores a full list of *all* encountered
' RGBA quads, and standard tracking would require 4 GB just for the tracking array.
'
'Upper bounds of tree nodes in lower levels of the tree automatically "float" based on encountered colors.
' This keeps memory usage especially low for images with normal (e.g. < 1 mln) unique pairs.  Leaf nodes
' are an exception to this rule, as they have the largest impact on resource demands, so they automatically
' "float" both lower and upper bounds, with specialized handling used for fully opaque pixels (as even
' true 32-bpp images tend to concentrate opaque pixels across a specific color range, so we don't always
' need to allocate leaf nodes if all occurrences of a given color are fully opaque).
'
'To further reduce memory consumption, the lower-bound of *all* nodes could be made dynamic as well,
' but this would impose performance penalties as we'd need to perform an additional bound-check on all
' accesses.  Similarly, upper-bounds of higher tree levels could also be made to float, but the resource
' gains from this are small as resource consumption directly correlates to tree depth - hence why the
' current arrangement is sort of a "sweet spot" between perf and resource constraints.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'A normal (elegant) solution to this problem would involve nested classes and simple recursive
' functions to add/count colors, but VB6 classes are heavyweight and extremely slow to allocate+free
' when using them in the 100,000+ object range.  As such, we cheat and use an array of dynamic types
' with tracked indices that mimic a b-tree arrangement.

'At present, we use a mathematically pleasing phi growth strategy.
' (https://stackoverflow.com/questions/1100311/what-is-the-ideal-growth-rate-for-a-dynamically-allocated-array)
Private Const GROWTH_STRATEGY As Double = ((1# + 2.23606797749979) * 0.5)

'By default, RGBA values are tracked.  Set this value to FALSE if you want to forcibly restrict
' tracking to RGB triplets only.
Private m_TrackRGBA As Boolean

Private Type PD_ColorEntry
    ceChildIndices() As Long    'Never initialized until an object is touched, and perhaps not even then;
End Type                        ' see the AddColor function for implementation details.

'Tree nodes are stored in a standard array to improve performance and de/construction time.
' Initial tree size is not especially relevant to either perf or resource usage, but 8k provides
' a nice baseline for "normal" images.
Private Const INIT_TABLE_SIZE As Long = 8192
Private m_NumOfEntries As Long
Private m_Colors() As PD_ColorEntry

'The actual color count of both RGB triplets and RGB quads; we calculate these as colors are added,
' so no second pass is required to traverse the "tree" and calculate final totals.
Private m_NumOfRGBTriplets As Long, m_NumOfRGBAQuads As Long

'Used for fast lookup indices
Private m_PowersOfTwo(0 To 31) As Long, m_DivBy32(0 To 255) As Byte

'Used as a persistent temporary buffer when changing the lower bound of existing arrays, since VB
' doesn't allow this w/ ReDim Preserve.
Private m_PersistentBuffer() As Byte

'Add a color to the running table.  Returns TRUE if the color is unique; FALSE if it already exists
Friend Function AddColor(ByVal srcR As Byte, ByVal srcG As Byte, ByVal srcB As Byte, ByVal srcA As Byte) As Boolean
    
    'If this is the first call, prep the first color entry (which handles red indices)
    If (m_NumOfEntries = 0) Then
        ReDim m_Colors(0 To INIT_TABLE_SIZE - 1) As PD_ColorEntry
        AllocateNode 255
        ReDim m_Colors(0).ceChildIndices(0 To 255) As Long
    End If
    
    'We now need to traverse child indices to find the correct leaf node.  This code is ugly
    ' but very fast.
    
    'First, see if a matching green node already exists.
    Dim gIndex As Long, bIndex As Long, aIndex As Long
    gIndex = m_Colors(0).ceChildIndices(srcR)
    If (gIndex = 0) Then
    
        'A green node does not exist.  (This will happen when adding a color with a novel
        ' green value for this particular red value.)  Allocate a new node now.
        ' (NOTE: for slower performance but lower memory usage, you can pass srcG as the
        ' initial UBound for the allocated array.  If you do this, you must also uncomment
        ' the Else block below - that will test the UBound on subsequent calls, and ensure
        ' that it is never exceeded.  (Because g indices appear at the 2nd level of the tree,
        ' they do not contribute nearly as much to resource usage as b indices, which appear
        ' at the 3rd level - and thus, there are potentially 256 of them for every 1 green
        ' node; as such, we use faster but resource-heavier allocations for green values.)
        'gIndex = AllocateNode(srcG)
        gIndex = AllocateNode(255)
        m_Colors(0).ceChildIndices(srcR) = gIndex
        
    'See above statement on green node allocation; for slightly slower but lower-resource
    ' tree construction, uncomment both this Else block and the special AllocateNode() call above.
    'Else
    '    If (srcG > UBound(m_Colors(gIndex).ceChildIndices)) Then ReDim Preserve m_Colors(gIndex).ceChildIndices(0 To srcG) As Long
    End If
    
    'Repeat the above steps, but this time for the green node (which points to a blue node).
    bIndex = m_Colors(gIndex).ceChildIndices(srcG)
    If (bIndex = 0) Then
    
        'A corresponding blue node does not exist for this green value.  Allocate one now.
        bIndex = AllocateNode(srcB)
        m_Colors(gIndex).ceChildIndices(srcG) = bIndex
    
    'A blue node already exists for this red+green pair.  Ensure its lookup table is large enough
    ' to address the current value.
    Else
        If (srcB > UBound(m_Colors(bIndex).ceChildIndices)) Then ReDim Preserve m_Colors(bIndex).ceChildIndices(0 To srcB) As Long
    End If
    
    'Repeat the above steps, but this time for the blue node (which points to an alpha node).
    aIndex = m_Colors(bIndex).ceChildIndices(srcB)
    If (aIndex = 0) Then
        
        'This is a unique RGB value; count it as part of the RGB triplet count
        m_NumOfRGBTriplets = m_NumOfRGBTriplets + 1
        
        'Alpha tracking is not mandatory; if you don't care about alpha values, you can set
        ' m_TrackRGBA to FALSE to forcibly track only unique RGB entries.
        If m_TrackRGBA Then
        
            'Because alpha values of 255 are most common, use a special flag to denote
            ' "255".  This spares us from allocating memory if a color only appears in pixels
            ' with full opacity.
            If (srcA = 255) Then
                aIndex = -1
                m_Colors(bIndex).ceChildIndices(srcB) = -1
                m_NumOfRGBAQuads = m_NumOfRGBAQuads + 1
                AddColor = True
                Exit Function
                
            'If alpha is not 255, allocate a leaf node to store specific alpha values.
            Else
            
                'A corresponding alpha node does not exist for this blue value.  Allocate one now,
                ' and note that we use a special "smaller" allocation (bit-flags only) to reduce
                ' memory requirements; this makes a huge difference on images with tons of variable
                ' opacity.
                aIndex = AllocateNode(m_DivBy32(srcA), True)
                m_Colors(bIndex).ceChildIndices(srcB) = aIndex
                
            End If
        
        'When not tracking RGBA quads, we still need to set the child index to an arbitrary non-zero
        ' value to avoid initializing it again.
        Else
            m_Colors(bIndex).ceChildIndices(srcB) = -1
            AddColor = True
            Exit Function
        End If
    
    End If
    
    'We now have a pointer to an "alpha" node, which completes this quadruplet.  We use the child
    ' array of this node to store "found" indicators only - if the current flag is 0, we know this
    ' is the first time this color quadruplet has been added to the list.
    '
    'Note also that this table uses a smaller index array (32 bytes only, instead of 1024 -
    ' just enough for 256 bit flags).  Each bit marks whether that opacity has been "found" for this
    ' particular triplet.  Note that this *greatly* reduces the overall memory footprint of the
    ' detection process.
    If m_TrackRGBA Then
        
        'Because opacities of 255 tend to be common (even in true RGBA images), we use a special
        ' flag for RGBA data.
        If (aIndex = -1) Then
        
            'If the current pixel also has an alpha value of 255, we don't need to do anything.
            If (srcA = 255) Then
                AddColor = False
                Exit Function
                
            'If the current pixel does *not* have an alpha value of 255, we need to allocate a full
            ' alpha leaf node.  Allocate it for the original alpha value of 255, make sure the 255
            ' value is flagged, then proceed as normal with adding this node.
            Else
                aIndex = AllocateNode(m_DivBy32(255), True)
                m_Colors(bIndex).ceChildIndices(srcB) = aIndex
                bIndex = m_DivBy32(255)
                m_Colors(aIndex).ceChildIndices(bIndex) = m_Colors(aIndex).ceChildIndices(bIndex) Or m_PowersOfTwo(255 And 31)
            End If
        
        End If
        
        'Both lower- and upper-bounds of leaf node lookups are allowed to "float" (e.g. they use
        ' only the lowest and highest values added to that node thus far).  This greatly reduces
        ' memory constraints as we don't allocate a full 32-bytes unless a particular RGB triplet
        ' is actually paired with alpha values stretching from 0 to 255
        Dim tstLBound As Long, tstUBound As Long
        tstLBound = LBound(m_Colors(aIndex).ceChildIndices)
        tstUBound = UBound(m_Colors(aIndex).ceChildIndices)
        bIndex = m_DivBy32(srcA)
        
        'Lower bound is out of range; we must perform a manual buffer resize
        If (bIndex < tstLBound) Then
        
            'Copy the existing entries to a safe temp buffer
            CopyMemoryStrict VarPtr(m_PersistentBuffer(0)), VarPtr(m_Colors(aIndex).ceChildIndices(tstLBound)), (tstUBound - tstLBound + 1) * 4
            
            'Resize the array
            ReDim m_Colors(aIndex).ceChildIndices(bIndex To tstUBound) As Long
            
            'Copy the data back
            CopyMemoryStrict VarPtr(m_Colors(aIndex).ceChildIndices(tstLBound)), VarPtr(m_PersistentBuffer(0)), (tstUBound - tstLBound + 1) * 4
        
        'Upper bound is out of range; normal ReDim Preserve will work
        ElseIf (bIndex > tstUBound) Then
            ReDim Preserve m_Colors(aIndex).ceChildIndices(tstLBound To bIndex) As Long
        End If
        
        'If the current flag bit for this alpha value is 0, we've encountered a new RGBA quad.
        ' Set the flag to 1 and increment our running RGBA unique quad count.
        If ((m_Colors(aIndex).ceChildIndices(bIndex) And m_PowersOfTwo(srcA And 31)) = 0) Then
            m_NumOfRGBAQuads = m_NumOfRGBAQuads + 1
            m_Colors(aIndex).ceChildIndices(bIndex) = m_Colors(aIndex).ceChildIndices(bIndex) Or m_PowersOfTwo(srcA And 31)
            AddColor = True
        End If
        
    End If
    
End Function

'Allocate a new node.  Returns the index of the allocated node.
Private Function AllocateNode(ByVal startValue As Long, Optional ByVal useSmallChildTable As Boolean = False) As Long
    If (m_NumOfEntries > UBound(m_Colors)) Then ReDim Preserve m_Colors(0 To m_NumOfEntries * GROWTH_STRATEGY - 1) As PD_ColorEntry
    With m_Colors(m_NumOfEntries)
        If useSmallChildTable Then
            ReDim .ceChildIndices(startValue To startValue) As Long
        Else
            ReDim .ceChildIndices(0 To startValue) As Long
        End If
    End With
    AllocateNode = m_NumOfEntries
    m_NumOfEntries = m_NumOfEntries + 1
End Function

'After counting colors, you can use this function to see if a specified color exists in the color tree.
Friend Function DoesColorExist(ByVal srcR As Long, ByVal srcG As Long, ByVal srcB As Long) As Boolean
    
    DoesColorExist = False
    
    If (Me.GetUniqueRGBACount > 0) Then
        
        If (m_Colors(0).ceChildIndices(srcR) <> 0) Then
            
            'R exists in the tree...
            Dim gIndex As Long
            gIndex = m_Colors(0).ceChildIndices(srcR)
            If (gIndex <> 0) And (srcG <= UBound(m_Colors(gIndex).ceChildIndices)) Then
                
                'G exists in the tree...
                Dim bIndex As Long
                bIndex = m_Colors(gIndex).ceChildIndices(srcG)
                If (bIndex <> 0) And (srcB <= UBound(m_Colors(bIndex).ceChildIndices)) Then
                    
                    'B exists in the tree...
                    Dim aIndex As Long
                    aIndex = m_Colors(bIndex).ceChildIndices(srcB)
                    DoesColorExist = (aIndex <> 0)
                
                End If
                
            End If
        
        End If
        
    End If

End Function

'After counting colors, you can use this function to retrieve a matching palette, e.g. a list of every
' color in the tree.  (While this works on any size color tree, this function is primarily designed
' for use on images with 256 colors or less.)
Friend Function GetPalette(ByRef dstPalette() As RGBQuad) As Boolean
    
    If (Me.GetUniqueRGBACount > 0) Then
        
        ReDim dstPalette(0 To Me.GetUniqueRGBACount - 1) As RGBQuad
        
        Dim r As Long, g As Long, b As Long, a As Long, cCount As Long
        Dim gIndex As Long, bIndex As Long, aIndex As Long
        cCount = 0
        
        'Iterate encountered r values
        For r = 0 To 255
            If (m_Colors(0).ceChildIndices(r) <> 0) Then
                gIndex = m_Colors(0).ceChildIndices(r)
                
                'Iterate encountered g values
                For g = 0 To UBound(m_Colors(gIndex).ceChildIndices)
                    If (m_Colors(gIndex).ceChildIndices(g) <> 0) Then
                        bIndex = m_Colors(gIndex).ceChildIndices(g)
                    
                        'Iterate encountered b values
                        For b = 0 To UBound(m_Colors(bIndex).ceChildIndices)
                            If (m_Colors(bIndex).ceChildIndices(b) <> 0) Then
                            
                                'If we're in RGBA mode, we need to iterate possible alpha values next
                                If m_TrackRGBA Then
                                
                                    aIndex = m_Colors(bIndex).ceChildIndices(b)
                                    
                                    '-1 is a special flag for RGB triplets with an opacity value of 255.
                                    ' If encountered, we don't need to loop alpha values as only full
                                    ' opacity pixels were encountered for this color.
                                    If (aIndex = -1) Then
                                        With dstPalette(cCount)
                                            .Red = r
                                            .Green = g
                                            .Blue = b
                                            .Alpha = 255
                                        End With
                                        cCount = cCount + 1
                                    
                                    'Alpha values other than 255 were encountered for this RGB triplet.
                                    ' Iterate through all of them, and return a new color for each
                                    ' flag bit.
                                    Else
                                        
                                        For a = LBound(m_Colors(aIndex).ceChildIndices) * 32 To UBound(m_Colors(aIndex).ceChildIndices) * 32 + 31
                                            
                                            'Alpha values use a truncated table; remap (a) into an index + flag bit
                                            If ((m_Colors(aIndex).ceChildIndices(m_DivBy32(a)) And m_PowersOfTwo(a And 31)) <> 0) Then
                                                With dstPalette(cCount)
                                                    .Red = r
                                                    .Green = g
                                                    .Blue = b
                                                    .Alpha = a
                                                End With
                                                cCount = cCount + 1
                                            End If
                                            
                                        Next a
                                        
                                    End If
                                    
                                'If we're in RGB mode, this is a unique triple; add it now and assume full opacity
                                Else
                                    With dstPalette(cCount)
                                        .Red = r
                                        .Green = g
                                        .Blue = b
                                        .Alpha = 255
                                    End With
                                    cCount = cCount + 1
                                End If
                            
                            End If
                        Next b
                    
                    End If
                Next g
                
            End If
            
        Next r
        
        GetPalette = True
    
    'A count hasn't been performed yet, so no color data exists
    Else
        GetPalette = False
    End If

End Function

'Return the number of unique RGB triplets for this branch
Friend Function GetUniqueRGBCount() As Long
    GetUniqueRGBCount = m_NumOfRGBTriplets
End Function

'Return the number of unique RGB quads for this branch
Friend Function GetUniqueRGBACount() As Long
    If m_TrackRGBA Then
        GetUniqueRGBACount = m_NumOfRGBAQuads
    Else
        GetUniqueRGBACount = m_NumOfRGBTriplets
    End If
End Function

'After counting colors, you can use this function to find an unused color in the current tree.
' (This is extremely helpful during PNG export, to try and find a good tRNS candidate.)
'
'Returns: TRUE if a valid color is found, FALSE if this function fails.  It will only fail if
' either the tree does not contain any colors yet (your fault), or if the tree contains all
' 16.7 million colors and there is no unused one to return (highly unlikely, but we're awesome
' so we do actually cover this case ;).
Friend Function GetUnusedColor(ByRef dstRed As Long, ByRef dstGreen As Long, ByRef dstBlue As Long) As Boolean
    
    If (Me.GetUniqueRGBACount > 0) Then
        
        GetUnusedColor = True
        
        Dim r As Long, g As Long, b As Long
        Dim gIndex As Long, bIndex As Long
        
        'Iterate encountered r values
        For r = 0 To 255
            
            'This red value is unused in the image.  Return any arbitrary green/blue value and exit.
            If (m_Colors(0).ceChildIndices(r) = 0) Then
                dstRed = r
                dstGreen = 0
                dstBlue = 0
                Exit Function
            
            'This red value is used, but maybe we can find a green/blue pair it isn't
            ' paired with.
            Else
                
                gIndex = m_Colors(0).ceChildIndices(r)
                For g = 0 To UBound(m_Colors(gIndex).ceChildIndices)
                    
                    'This green value is unused in the image.  Return any arbitrary blue value and exit.
                    If (m_Colors(gIndex).ceChildIndices(g) = 0) Then
                        dstRed = r
                        dstGreen = g
                        dstBlue = 0
                        Exit Function
                    
                    'This green value is used, but maybe we can find a blue value it isn't
                    ' paired with.
                    Else
                    
                        bIndex = m_Colors(gIndex).ceChildIndices(g)
                        For b = 0 To UBound(m_Colors(bIndex).ceChildIndices)
                            
                            'This blue value is unused in the image.  Return it and exit.
                            If (m_Colors(bIndex).ceChildIndices(b) = 0) Then
                                dstRed = r
                                dstGreen = g
                                dstBlue = b
                                Exit Function
                            End If
                            
                        Next b
                    
                    End If
                    
                Next g
                
            End If
            
        Next r
        
        'If by some miracle we reach this point, it means every single RGB triple is used in this image
        ' (so the image must have at least 16.7 million pixels - interesting?).  Return false.
        GetUnusedColor = False
    
    'A count hasn't been performed yet, so no color data exists
    Else
        GetUnusedColor = False
    End If

End Function

Friend Sub SetAlphaTracking(ByVal newState As Boolean)
    m_TrackRGBA = newState
End Sub

Private Sub Class_Initialize()
    
    'By default, RGBA data is always analyzed
    m_TrackRGBA = True
    
    'Prep some persistent luts
    Dim i As Long
    For i = 0 To 30
        m_PowersOfTwo(i) = 2 ^ i
    Next i
    m_PowersOfTwo(31) = &H80000000
    
    For i = 0 To 255
        m_DivBy32(i) = i \ 32
    Next i
    
    'Prep a persistent buffer for holding data while resizing arrays
    ReDim m_PersistentBuffer(0 To 31) As Byte
    
End Sub
