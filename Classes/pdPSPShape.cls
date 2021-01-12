VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPSPShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PSP (PaintShop Pro) Vector Shape Container
'Copyright 2020-2021 by Tanner Helland
'Created: 31/December/20
'Last updated: 11/January/21
'Last update: polyline-type shapes are now parsed correctly!  This is a big victory.
'
'This class describes a single vector "shape" inside a JASC/Corel Paint Shop Pro image file.
' It has been custom-built for PhotoDemon, with a special emphasis on parsing performance.
' Vector rendering elements are automatically translated to pd2D as needed (and pd2D ultimately
' handles all rendering!).
'
'Unless otherwise noted, all code in this class is my original work.  I've based my work off the
' "official" PSP spec at this URL (link good as of December 2020):
' ftp://ftp.corel.com/pub/documentation/PSP/
'
'Older PSP specs were also useful.  You may be able to find them here (link good as of December 2020);
' look for files with names like "psp8spec.pdf":
' http://www.telegraphics.com.au/svn/pspformat/trunk
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Some shape elements are "blocks".  These use the same magic number as all other PSP blocks.
Private Const PSP_BLOCK_MARKER As Long = &H4B427E

'/* Vector shape types.  */
Private Enum PSPVectorShapeType
    keVSTUnknown = 0    '// Undefined vector type
    keVSTText           '// Shape represents lines of text
    keVSTPolyline       '// Shape represents a multiple segment line
    keVSTEllipse        '// Shape represents an ellipse (or circle)
    keVSTPolygon        '// Shape represents a closed polygon
    keVSTGroup          '// Shape represents a group shape
End Enum

#If False Then
    Private Const keVSTUnknown = 0, keVSTText = 1, keVSTPolyline = 2, keVSTEllipse = 3, keVSTPolygon = 4, keVSTGroup = 5
#End If

'/* Shape property flags  */
Private Enum PSPShapeProperties
    keShapeAntiAliased = &H1        '// Shape is anti-aliased
    keShapeSelected = &H2           '// Shape is selected
    keShapeVisible = &H4            '// Shape is visible
End Enum

#If False Then
    Private Const keShapeAntiAliased = &H1, keShapeSelected = &H2, keShapeVisible = &H4
#End If

'/* Polyline node type flags.  */
Private Enum PSPPolylineNodeTypes
    keNodeUnconstrained = &H0       '// Default node type
    keNodeSmooth = &H1              '// Node is smooth
    keNodeSymmetric = &H2           '// Node is symmetric
    keNodeAligned = &H4             '// Node is aligned
    keNodeActive = &H8              '// Node is active
    keNodeLocked = &H10             '// Node is locked
    keNodeSelected = &H20           '// Node is selected
    keNodeVisible = &H40            '// Node is visible
    keNodeClosed = &H80             '// Node is closed
End Enum

#If False Then
    Private Const keNodeUnconstrained = &H0, keNodeSmooth = &H1, keNodeSymmetric = &H2, keNodeAligned = &H4, keNodeActive = &H8, keNodeLocked = &H10, keNodeSelected = &H20, keNodeVisible = &H40, keNodeClosed = &H80
#End If

'/* Paint style types.  */
Private Enum PSPPaintStyleType
    keStyleNone = &H0           '// No paint style info applies
    keStyleColor = &H1          '// Color paint style info
    keStyleGradient = &H2       '// Gradient paint style info
    keStylePattern = &H4        '// Pattern paint style info
    keStylePaper = &H8          '// Paper paint style info
    keStylePen = &H10           '// Organic pen paint style info
End Enum

#If False Then
    Private Const keStyleNone = &H0, keStyleColor = &H1, keStyleGradient = &H2, keStylePattern = &H4, keStylePaper = &H8, keStylePen = &H10
#End If

'/* Gradient type.  */
Private Enum PSPStyleGradientType
    keSGTLinear = 0     '// Linear gradient type
    keSGTRadial         '// Radial gradient type
    keSGTRectangular    '// Rectangular gradient type
    keSGTSunburst       '// Sunburst gradient type
End Enum

#If False Then
    Private Const keSGTLinear = 0, keSGTRadial = 1, keSGTRectangular = 2, keSGTSunburst = 3
#End If

'/* Paint Style Cap Type (Start & End).  */
Private Enum PSPStyleCapType
    keSCTCapFlat = 0        '// Flat cap type (was round in psp6)
    keSCTCapRound           '// Round cap type (was square in psp6)
    keSCTCapSquare          '// Square cap type (was flat in psp6)
    keSCTCapArrow           '// Arrow cap type
    keSCTCapCadArrow        '// Cad arrow cap type
    keSCTCapCurvedTipArrow  '// Curved tip arrow cap type
    keSCTCapRingBaseArrow   '// Ring base arrow cap type
    keSCTCapFluerDelis      '// Fluer de Lis cap type
    keSCTCapFootball        '// Football cap type
    keSCTCapXr71Arrow       '// Xr71 arrow cap type
    keSCTCapLilly           '// Lilly cap type
    keSCTCapPinapple        '// Pinapple cap type
    keSCTCapBall            '// Ball cap type
    keSCTCapTulip           '// Tulip cap type
End Enum

#If False Then
    Private Const keSCTCapFlat = 0, keSCTCapRound = 1, keSCTCapSquare = 2, keSCTCapArrow = 3, keSCTCapCadArrow = 4, keSCTCapCurvedTipArrow = 5, keSCTCapRingBaseArrow = 6, keSCTCapFluerDelis = 7, keSCTCapFootball = 8, keSCTCapXr71Arrow = 9
    Private Const keSCTCapLilly = 10, keSCTCapPinapple = 11, keSCTCapBall = 12, keSCTCapTulip = 13
#End If

'/* Paint Style Join Type.  */
Private Enum PSPStyleJoinType
    keSJTJoinMiter = 0      '// Miter join type
    keSJTJoinRound          '// Round join type
    keSJTJoinBevel          '// Bevel join type
End Enum

#If False Then
    Private Const keSJTJoinMiter = 0, keSJTJoinRound = 1, keSJTJoinBevel = 2
#End If

'/* Organic pen type.  */
Private Enum PSPStylePenType
    keSPTOrganicPenNone = 0     '// Undefined pen type
    keSPTOrganicPenMesh         '// Mesh pen type
    keSPTOrganicPenSand         '// Sand pen type
    keSPTOrganicPenCurlicues    '// Curlicues pen type
    keSPTOrganicPenRays         '// Rays pen type
    keSPTOrganicPenRipple       '// Ripple pen type
    keSPTOrganicPenWave         '// Wave pen type
    keSPTOrganicPen             '// Generic pen type
End Enum

#If False Then
    Private Const keSPTOrganicPenNone = 0, keSPTOrganicPenMesh = 1, keSPTOrganicPenSand = 2, keSPTOrganicPenCurlicues = 3, keSPTOrganicPenRays = 4, keSPTOrganicPenRipple = 5, keSPTOrganicPenWave = 6, keSPTOrganicPen = 7
#End If

'/* Text element types. */
Private Enum PSPTextElementType
    keTextElemUnknown = 0       '// Undefined text element type
    keTextElemChar              '// A single character code
    keTextElemCharStyle         '// A character style change
    keTextElemLineStyle         '// A line style change
End Enum

#If False Then
    Private Const keTextElemUnknown = 0, keTextElemChar = 1, keTextElemCharStyle = 2, keTextElemLineStyle = 3
#End If

'/* Text alignment types.  */
Private Enum PSPTextAlignment
    keTextAlignmentLeft = 0     '// Left text alignment
    keTextAlignmentCenter       '// Center text alignment
    keTextAlignmentRight        '// Right text alignment
End Enum

#If False Then
    Private Const keTextAlignmentLeft = 0, keTextAlignmentCenter = 1, keTextAlignmentRight = 2
#End If

'/* Text antialias modes.  */
Private Enum PSPAntialiasMode
    keNoAntialias = 0   '// Antialias off
    keSharpAntialias    '// Sharp
    keSmoothAntialias   '// Smooth
End Enum

#If False Then
    Private Const keNoAntialias = 0, keSharpAntialias = 1, keSmoothAntialias = 2
#End If

'/* Text flow types  */
Private Enum PSPTextFlow
    keTFHorizontalDown = 0  '// Horizontal then down
    keTFVerticalLeft        '// Vertical then left
    keTFVerticalRight       '// Vertical then right
    keTFHorizontalUp        '// Horizontal then up
End Enum

#If False Then
    Private Const keTFHorizontalDown = 0, keTFVerticalLeft = 1, keTFVerticalRight = 2, keTFHorizontalUp = 3
#End If

'/* Character style flags.  */
Private Enum PSPCharacterProperties
    keStyleItalic = &H1         '// Italic property bit
    keStyleStruck = &H2         '// Strike-out property bit
    keStyleUnderlined = &H4     '// Underlined property bit
    keStyleWarped = &H8         '// Warped property bit
    keStyleAntiAliased = &H10   '// Anti-aliased property bit
End Enum

#If False Then
    Private Const keStyleItalic = &H1, keStyleStruck = &H2, keStyleUnderlined = &H4, keStyleWarped = &H8, keStyleAntiAliased = &H10
#End If

'TRUE if the shape was loaded and validated successfully
Private m_ShapeOK As Boolean

'All vector shapes start with the same uniform header.  Their parsing then branches based on the
' shape's type.
Private Type PSP_ShapeAttributes
    sa_Size As Long                 'DWORD - length of vector shape attributes chunk
    sa_Name As String               'WORD/[char] - variable length string chunk
    sa_Type As PSPVectorShapeType   'WORD - type of vector shape (must be one of PSPVectorShapeType)
    sa_Flags As PSPShapeProperties  'DWORD - series of property flags (in PSPShapeProperties), like visibility or AA
    sa_UniqueID As Long             'DWORD - unique ID (within the layer), 1-based (0 is invalid)
    sa_LinkedShapeID As Long        'DWORD - ID of a linked shape, if any (used for text on path); 0 means no linked shape
    'Future Expansion fields (skip using initial size)
End Type

Private m_ShapeAttributes As PSP_ShapeAttributes

'Beyond the shape attributes struct, shape contents vary by shape type.  (For example,
' a text layer uses a totally different structure than a polygon.)

'Polylines, ellipses, and polygons use the same format: an attribute header, a definition header
' (describing node count, basically), then a list of nodes, including control points.
Private Type PSP_PolylineAttributes
    pa_Size As Long
    pa_Stroked As Boolean
    pa_Filled As Boolean
    pa_StyledLine As Boolean
    pa_StrokeWidth As Double
    pa_StartCapType As PSPStyleCapType
    pa_StartCapMultiplier As Boolean
    pa_StartCapWidthMultiplier As Double
    pa_StartCapHeightMultiplier As Double
    pa_EndCapType As PSPStyleCapType
    pa_EndCapMultiplier As Boolean
    pa_EndCapWidthMultiplier As Double
    pa_EndCapHeightMultiplier As Double
    pa_LineJoin As PSPStyleJoinType
    pa_MiterLimit As Double
    'Future expansion fields possible; use Size member to skip
End Type

Private m_PolylineAttributes As PSP_PolylineAttributes

'When painting vector shapes, PSP files don't use normal "stroke" or "fill" terminology.
' Instead, pretty much everything is treated as a "fill", and "fills" are described by
' "paint style" blocks.  (These can describe solid colors, gradients, patterns, and more.)
'
'All vector shapes can embed three "paint style" blocks: one for outline (stroke), one for
' fill, and a 3rd PSP-specific one calle "styled line" (haven't studied this yet).
'
'Instead of storing PSP-specific structs with all that data, this class tries to produce
' three pd2D brushes instead.  Note that pd2D brushes do not always behave identically to
' PSP ones; this is difficult to rectify, but I'll always try to match up inconsistencies
' when people find them.
Private m_BrushStroke As pd2DBrush, m_BrushFill As pd2DBrush, m_PenStyled As pd2DPen

'Polyline vectors are comprised of a list of nodes
Private Type PSP_PolylineNode
    pn_Size As Long
    pn_X As Double
    pn_Y As Double
    pn_hX1 As Double
    pn_hY1 As Double
    pn_hX2 As Double
    pn_hY2 As Double
    pn_MoveTo As Boolean
    pn_NodeFlags As PSPPolylineNodeTypes
    'Spec allows for expansion in the future
End Type

Private m_NodeCount As Long, m_Nodes() As PSP_PolylineNode

'Polyline data is ultimately converted to a pd2D path object
Private m_PolylineShape As pd2DPath

'Total block length and initial block offset should always be used to realign the stream pointer
' before exiting this class (under *any* circumstances, including failure states).
Private m_BlockOffset As Long, m_TotalLength As Long

'Files from PSP 6 use a totally different struct.  Because the PSP spec is ill-conceived, there is
' no natural way to identify these files (e.g. a header size or similar).  Instead, we can only rely
' on the header of *the entire file* to try and pick out these files if they occur.
Private m_FileIsPSPv6 As Boolean

Friend Function IsShapeOK() As Boolean
    IsShapeOK = m_ShapeOK
End Function

'Assuming the source stream is pointing at the start of a vector shape block, attempt to load the shape.
' Returns psp_Success if successful, psp_Warning if stream alignment is okay but shape data is not,
' psp_Failure if stream alignment is unsaveable.  (If psp_Failure is returned, check initial stream
' pointer alignment - it may not have been pointing at a shape block when you called this function!)
'
'IMPORTANTLY: on psp_Success or psp_Warning, the passed stream pointer will now point at the *end* of
' this block.  You can simply continue reading the file as-is.  On failure, however, stream position
' is *not* guaranteed (mostly because if initial block validation fails, we have no way to reorient the
' pointer in a meaningful way - we can only reset it).  On failure, you need to abandon further parsing.
Friend Function LoadShape(ByRef srcStream As pdStream, ByRef srcWarnings As pdStringStack, ByRef srcHeader As PSPImageHeader) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "LoadShape"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'This shape will only be marked "OK" if we can retrieve at least one valid channel for it
    m_ShapeOK = False
    
    'The caller should have already performed block validation, but we need to be extra careful
    ' because all subsequent stream alignment depends on this.
    Const PSP_BLOCK_MARKER As Long = &H4B427E
    If (srcStream.ReadLong() <> PSP_BLOCK_MARKER) Then
        LoadShape = psp_Failure
        InternalError funcName, "stream misaligned", srcWarnings
        Exit Function
    End If
    
    Dim blockID As PSPBlockID
    blockID = srcStream.ReadIntUnsigned()
    If (blockID <> PSP_SHAPE_BLOCK) Then
        LoadShape = psp_Failure
        InternalError funcName, "not a shape block: " & blockID, srcWarnings
        Exit Function
    End If
    
    'Thankfully, PSPv5 files are not a concern here (they didn't support vector layers),
    ' so we only need to deal with v6+ block format
    m_TotalLength = srcStream.ReadLong()
    m_BlockOffset = srcStream.GetPosition()
    
    'PSP v6 files use a totally different layout, and due to the asinine design of these blocks
    ' (where header size includes a ton of huge, variable-length blocks that follow instead of
    ' just defining the damn header), it's impossible to detect as-we-go.  Instead, we need to
    ' rely on the master version at the head of the file to tell us how to parse.
    m_FileIsPSPv6 = (srcHeader.psph_VersionMajor <= 4)
    
    'Immediately following the block header is a "shape attributes" chunk.  This tells us
    ' the shape name and type, which lead to different blueprints for parsing the remainder
    ' of this shape block.
    With m_ShapeAttributes
        .sa_Size = srcStream.ReadLong()
        If (.sa_Size <= 0) Then
            InternalError funcName, "bad shape attribute size: " & .sa_Size, srcWarnings
            okToProceed = psp_Failure
            GoTo EarlyTermination
        End If
        
        'v6 files do not supply a shape name
        If (Not m_FileIsPSPv6) Then
        
            'In PSP files, all variable-length strings are prefixed by a WORD length
            Dim sLen As Long
            sLen = srcStream.ReadIntUnsigned()
            .sa_Name = srcStream.ReadString_UTF8(sLen)
        
        End If
        
        .sa_Type = srcStream.ReadIntUnsigned()
        .sa_Flags = srcStream.ReadLong()
        .sa_UniqueID = srcStream.ReadLong()
        .sa_LinkedShapeID = srcStream.ReadLong()
        
    End With
    
    'Use the attribute's chunk size to realign the pointer
    srcStream.SetPosition m_BlockOffset + m_ShapeAttributes.sa_Size, FILE_BEGIN
    
    'What we're pointing out now varies by shape type (above)
    Select Case m_ShapeAttributes.sa_Type
    
        Case keVSTText
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Found shape: text"
            'Do text stuff
            
        Case keVSTPolyline, keVSTEllipse, keVSTPolygon
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Found shape: polyline"
            okToProceed = ReadPolylines(srcStream, srcWarnings, srcHeader)
            m_ShapeOK = (okToProceed = psp_Success)
            'These all use polyline definitions
        
        Case keVSTGroup
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Found shape: vector group marker"
            'Like raster layer groups, this indicator simply reports the number of shapes
            ' that belong in this group.  The spec is very unclear here, just stating:
            ' "Number of shapes in this group (i.e., number of following Vector Shape
            '  Sub-Blocks belonging to this group)."  More investigation needed!
    
    End Select
    
    'Our work here is done!
    LoadShape = okToProceed
    
    'Before exiting, move the stream pointer to the end of this block.
EarlyTermination:
    srcStream.SetPosition m_BlockOffset + m_TotalLength, FILE_BEGIN
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description, srcWarnings
    srcWarnings.AddString "Internal error in pdPSPShape." & funcName & ", #" & Err.Number & ": " & Err.Description
    LoadShape = psp_Failure
    
End Function

'"Polyline" shapes include polyline, ellipse, and polygon definitions
Private Function ReadPolylines(ByRef srcStream As pdStream, ByRef srcWarnings As pdStringStack, ByRef srcHeader As PSPImageHeader) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "ReadPolylines"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    Dim origStreamPos As Long
    origStreamPos = srcStream.GetPosition()
    
    'This step starts out pretty simple - just fill the attribute header!
    With m_PolylineAttributes
        
        .pa_Size = srcStream.ReadLong()
        If (.pa_Size <= 0) Then
            InternalError funcName, "bad polyline attribute size: " & .pa_Size, srcWarnings
            okToProceed = psp_Failure
            GoTo EarlyTermination
        End If
        
        'v6 files place the shape name *here*
        If m_FileIsPSPv6 Then
            Dim sLen As Long
            sLen = srcStream.ReadIntUnsigned()
            m_ShapeAttributes.sa_Name = srcStream.ReadString_UTF8(sLen)
        End If
        
        .pa_Stroked = (srcStream.ReadByte() <> 0)
        
        'For the rest of this struct, there are critical variations in v6 PSP files.
        ' I won't comment all of them, but unfortunately there's no easy way to handle
        ' their vastly different struct signatures.
        If m_FileIsPSPv6 Then
            .pa_StrokeWidth = srcStream.ReadIntUnsigned()
            .pa_Filled = (srcStream.ReadByte() <> 0)
            .pa_StartCapType = srcStream.ReadByte()
            .pa_EndCapType = .pa_StartCapType
            .pa_LineJoin = srcStream.ReadByte()
            .pa_MiterLimit = srcStream.ReadDouble()
        Else
            .pa_Filled = (srcStream.ReadByte() <> 0)
            .pa_StyledLine = (srcStream.ReadByte() <> 0)
            .pa_StrokeWidth = srcStream.ReadDouble()
            .pa_StartCapType = srcStream.ReadByte()
            .pa_StartCapMultiplier = (srcStream.ReadByte() <> 0)
            .pa_StartCapWidthMultiplier = srcStream.ReadDouble()
            .pa_StartCapHeightMultiplier = srcStream.ReadDouble()
            .pa_EndCapType = srcStream.ReadByte()
            .pa_EndCapMultiplier = (srcStream.ReadByte() <> 0)
            .pa_EndCapWidthMultiplier = srcStream.ReadDouble()
            .pa_EndCapHeightMultiplier = srcStream.ReadDouble()
            .pa_LineJoin = srcStream.ReadByte()
            .pa_MiterLimit = srcStream.ReadDouble()
        End If
            
    End With
    
    'With the header read, use its size parameter to skip to the end of the header
    srcStream.SetPosition origStreamPos + m_PolylineAttributes.pa_Size, FILE_BEGIN
    
    'Three style blocks now follow before we reach the actual polyline definition
    ' for this shape: one for stroke style, one for fill style, and one for special
    ' "styled line" stroking (dashes, etc).  A separate function attempts to produce
    ' corresponding pd2D objects for each of these style types.
    '
    '(Note that these functions manage their own stream alignment.)
    If (okToProceed < psp_Failure) Then okToProceed = BuildBrushForPaintBlock(srcStream, srcWarnings, m_BrushStroke, Nothing)
    If (okToProceed < psp_Failure) Then okToProceed = BuildBrushForPaintBlock(srcStream, srcWarnings, m_BrushFill, Nothing)
    
    'v6 files do not include the 3rd style block (for non-solid lines, e.g. dashed/dotted lines)
    If (Not m_FileIsPSPv6) Then
        If (okToProceed < psp_Failure) Then okToProceed = BuildBrushForPaintBlock(srcStream, srcWarnings, Nothing, m_PenStyled)
    End If
    
    'With any pens and/or brushes constructed, we can now move onto reading in the actual PolyLine.
    Dim chunkSize As Long, origPosition As Long
    origPosition = srcStream.GetPosition()
    chunkSize = srcStream.ReadLong()
    If (chunkSize > 0) Then
        
        m_NodeCount = srcStream.ReadLong()
        
        If (m_NodeCount > 0) Then
            
            ReDim m_Nodes(0 To m_NodeCount - 1) As PSP_PolylineNode
            
            'Use chunk size to skip potential expansion fields
            srcStream.SetPosition origPosition + chunkSize, FILE_BEGIN
            
            'Iterate all nodes.  Note that nodes also support future expansion fields,
            ' so the stream must be manually aligned after each read.
            Dim i As Long
            For i = 0 To m_NodeCount - 1
                
                origPosition = srcStream.GetPosition()
                
                With m_Nodes(i)
                    .pn_Size = srcStream.ReadLong()
                    .pn_X = srcStream.ReadDouble()
                    .pn_Y = srcStream.ReadDouble()
                    .pn_hX1 = srcStream.ReadDouble()
                    .pn_hY1 = srcStream.ReadDouble()
                    .pn_hX2 = srcStream.ReadDouble()
                    .pn_hY2 = srcStream.ReadDouble()
                    .pn_MoveTo = (srcStream.ReadByte() <> 0)
                    .pn_NodeFlags = srcStream.ReadIntUnsigned()
                End With
                
                srcStream.SetPosition origPosition + m_Nodes(i).pn_Size, FILE_BEGIN
                
            Next i
            
            'With the polyline fully read, it's now time to construct a matching pd2D path object
            okToProceed = BuildPD2DPolygon(srcWarnings, srcHeader)
            
        Else
            InternalError funcName, "no nodes", srcWarnings
        End If
            
    Else
        InternalError funcName, "polyline shape definition chunk size invalid", srcWarnings
        okToProceed = psp_Warning
        GoTo EarlyTermination
    End If
    
    'Unlike most other read functions in PD's PSP-centric collection, this reader does
    ' *not* align the stream before exiting.  That is up to our parent class due to the
    ' way PSP embeds chunk vs block sizes (this is a chunk reader, not a block reader).
EarlyTermination:
    ReadPolylines = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description, srcWarnings
    srcWarnings.AddString "Internal error in pdPSPShape." & funcName & ", #" & Err.Number & ": " & Err.Description
    ReadPolylines = psp_Failure
    
End Function

Friend Function RenderShapeIntoDIB(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad, ByVal srcPaletteSize As Long, ByRef srcHeader As PSPImageHeader) As Boolean
    
    Const funcName As String = "RenderShapeIntoDIB"
    
    'Failsafe checks only
    If (dstDIB Is Nothing) Then Exit Function
    
    'Palette usage is TODO
    
    'If this shape is "invisible", do nothing.  (This may change in the future if we place
    ' each shape onto its own layer FYI)
    If ((m_ShapeAttributes.sa_Flags And keShapeVisible) = 0) Then
        RenderShapeIntoDIB = True
        Exit Function
    End If
    
    'pd2D is used for rendering
    Dim dstSurface As pd2DSurface
    Set dstSurface = New pd2DSurface
    dstSurface.WrapSurfaceAroundPDDIB dstDIB
    
    'Antialiasing is stored in the parent shape attributes flags
    If ((m_ShapeAttributes.sa_Flags And keShapeAntiAliased) <> 0) Then
        dstSurface.SetSurfaceAntialiasing P2_AA_HighQuality
    Else
        dstSurface.SetSurfaceAntialiasing P2_AA_None
    End If
    
    'Half-offset pixels tend to look better in GDI+ rendering, but standard pixel offsets produce
    ' a result closer to PSPs, so we defer to it.
    dstSurface.SetSurfacePixelOffset P2_PO_Normal
     
    'Fill if flagged
    If m_PolylineAttributes.pa_Filled Then
    
        'Make sure our fill brush exists
        If (Not m_BrushFill Is Nothing) Then
            PD2D.FillPath dstSurface, m_BrushFill, m_PolylineShape
            RenderShapeIntoDIB = True
        Else
            InternalError funcName, "no fill brush", Nothing
        End If
        
        'TESTING ONLY: use this to stroke the outline of the path
        'Dim cPen As pd2DPen
        'Set cPen = New pd2DPen
        'cPen.SetPenColor vbBlack
        'cPen.SetPenWidth 1
        'PD2D.DrawPath dstSurface, cPen, m_PolylineShape
    
    'Stroke if flagged
    ElseIf m_PolylineAttributes.pa_Stroked Then
        
        If (Not m_BrushStroke Is Nothing) Then
        
            Dim shapePen As pd2DPen
            Set shapePen = New pd2DPen
            
            'If a special "styled" pen exists, use it as our base pen
            If (Not m_PenStyled Is Nothing) Then
                shapePen.ClonePen m_PenStyled
            Else
                shapePen.SetPenStartCap GetPDLineCapFromPSPLineCap(m_PolylineAttributes.pa_StartCapType)
                shapePen.SetPenEndCap GetPDLineCapFromPSPLineCap(m_PolylineAttributes.pa_EndCapType)
                'Cap multipliers are TODO
            End If
            
            'Apply separate pen settings from the *parent* shape object
            shapePen.SetPenWidth m_PolylineAttributes.pa_StrokeWidth
            
            Dim pdLineJoin As PD_2D_LineJoin
            Select Case m_PolylineAttributes.pa_LineJoin
                Case keSJTJoinMiter
                    pdLineJoin = P2_LJ_Miter
                Case keSJTJoinRound
                    pdLineJoin = P2_LJ_Round
                Case keSJTJoinBevel
                    pdLineJoin = P2_LJ_Miter
            End Select
            shapePen.SetPenLineJoin pdLineJoin
            
            shapePen.SetPenMiterLimit m_PolylineAttributes.pa_MiterLimit
            
            'Create a pen from the stroking brush specified in the file.
            ' (Note that gradient pens are TODO; they will require us to calculate a boundary somehow)
            shapePen.CreatePenFromBrush m_BrushStroke
            
            'Finally, perform the damn stroke!
            PD2D.DrawPath dstSurface, shapePen, m_PolylineShape
            
            RenderShapeIntoDIB = True
        
        Else
            InternalError funcName, "no stroke brush", Nothing
        End If
        
    End If
    
End Function

Private Function BuildBrushForPaintBlock(ByRef srcStream As pdStream, ByRef srcWarnings As pdStringStack, ByRef dstBrush As pd2DBrush, ByRef dstPen As pd2DPen) As PD_PSPResult

    On Error GoTo InternalVBError
    Const funcName As String = "BuildBrushForPaintBlock"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    Dim origStreamPos As Long
    origStreamPos = srcStream.GetPosition()
    
    'Always start with block ID verification.  If this fails, we importantly need to
    ' exit immediately as stream alignment is borked.
    Dim blockHeaderID As Long
    blockHeaderID = srcStream.ReadLong()
    If (blockHeaderID <> PSP_BLOCK_MARKER) Then
        InternalError funcName, "not block-aligned: " & blockHeaderID, srcWarnings
        BuildBrushForPaintBlock = psp_Failure
        Exit Function
    End If
    
    'This can be either a paint style or a line style entry.  Paint styles produce brushes,
    ' line styles produce pens.
    Dim blockID As PSPBlockID, blockLength As Long, endOfBlockHeader As Long
    blockID = srcStream.ReadIntUnsigned()
    blockLength = srcStream.ReadLong()
    endOfBlockHeader = srcStream.GetPosition()
    
    Dim chunkSize As Long
    
    'Build a brush
    If (blockID = PSP_PAINTSTYLE_BLOCK) Then
        
        Set dstBrush = New pd2DBrush
        
        'DEBUG ONLY: prep an opaque, white brush as a fallback in case other methods fail
        dstBrush.SetBrushColor vbWhite
        dstBrush.SetBrushOpacity 100!
        
        chunkSize = srcStream.ReadLong()
        If (chunkSize > 0) Then
        
            Dim paintStyleType As PSPPaintStyleType
            paintStyleType = srcStream.ReadIntUnsigned()
            
            'Before handling paint style type, jump to the end of this chunk.
            ' (The spec allows future expansion bytes here, and we must account for them.)
            srcStream.SetPosition endOfBlockHeader + chunkSize, FILE_BEGIN
            
            'Based on paint style type, we'll branch into a new parser.  Note that multiple
            ' brushes can be defined!  This allows e.g. a pattern to be mixed with a gradient.
            If (paintStyleType = keStyleNone) Then
                InternalError funcName, "no paint styles defined", srcWarnings
                BuildBrushForPaintBlock = psp_Warning
            Else
                
                If (paintStyleType And keStyleColor) Then
                    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "solid brush found; retrieving now..."
                    BuildBrushForPaintBlock = BuildBrush_Color(srcStream, srcWarnings, dstBrush)
                End If
                
                If (paintStyleType And keStyleGradient) Then
                    InternalError funcName, "gradient styles are TODO", srcWarnings
                    BuildBrushForPaintBlock = psp_Warning
                End If
                
                If (paintStyleType And keStylePattern) Then
                    InternalError funcName, "pattern styles are TODO", srcWarnings
                    BuildBrushForPaintBlock = psp_Warning
                End If
                
                If (paintStyleType And keStylePaper) Then
                    InternalError funcName, "paper styles are TODO", srcWarnings
                    BuildBrushForPaintBlock = psp_Warning
                End If
                
                If (paintStyleType And keStylePen) Then
                    InternalError funcName, "organic pen styles are TODO", srcWarnings
                    BuildBrushForPaintBlock = psp_Warning
                End If
                
            End If
        
        Else
            InternalError funcName, "bad brush-style chunk size: " & chunkSize, srcWarnings
            okToProceed = psp_Warning
            GoTo EarlyTermination
        End If
        
    'Build a pen.  From the spec:
    ' "The Line Style Sub-Block contains styled line data used by the outline paint style of
    ' vector shapes in the Vector Shape Sub-Block. As illustrated below, the Line Style Sub-Block
    ' consists of the Line Style Block Header, the Line Style Information Chunk, and the Line
    ' Style Entries Chunk."
    '
    'It's a little confusing because the line style block defines line caps and dash behavior
    ' (theoretically - this information is also in the vector shape header, so I have no idea
    ' which takes precedence).  The constructed line, however, is then filled with a brush
    ' defined by the preceding "paint style" block.  This is different from how most graphics
    ' libraries define "stroking" and "filling", with PSP treating all "strokes" as "fills".
    ' pd2D can do this, but it complicates things.
    ElseIf (blockID = PSP_LINESTYLE_BLOCK) Then
        
        'Line-style blocks are rarer but much easier to parse.  They're just basic pen data,
        ' with optional dashed-line definitions.
        Set dstPen = New pd2DPen
        
        chunkSize = srcStream.ReadLong()
        If (chunkSize > 0) Then
        
            'First come start cap and end cap properties
            Dim startCapType As PSPStyleCapType, useStartCapMultiplier As Boolean
            startCapType = srcStream.ReadByte()
            useStartCapMultiplier = (srcStream.ReadByte() <> 0)
            
            Dim startCapMultiplierWidth As Double, startCapMultiplierHeight As Double
            startCapMultiplierWidth = srcStream.ReadDouble()
            startCapMultiplierHeight = srcStream.ReadDouble()
                        
            Dim endCapType As PSPStyleCapType, useendCapMultiplier As Boolean
            endCapType = srcStream.ReadByte()
            useendCapMultiplier = (srcStream.ReadByte() <> 0)
            
            Dim endCapMultiplierWidth As Double, endCapMultiplierHeight As Double
            endCapMultiplierWidth = srcStream.ReadDouble()
            endCapMultiplierHeight = srcStream.ReadDouble()
            
            'From the spec: "TRUE if the line segment caps are linked to the shape's stroke
            ' outline caps, FALSE otherwise."
            Dim linkCapsFlag As Boolean
            linkCapsFlag = (srcStream.ReadByte() <> 0)
            
            'From the spec: "Number of dash-gap entries in the following Line Style Entries Chunk."
            Dim numDashGaps As Long
            numDashGaps = srcStream.ReadLong()
            
            'Expansion fields are possible in the future.  Use chunk size to align the stream.
            srcStream.SetPosition endOfBlockHeader + chunkSize, FILE_BEGIN
            
            'Retrieve [numDashGaps] entries (interpretation TBD)
            If (numDashGaps > 0) Then
                
                Dim lineDashGaps() As Long
                ReDim lineDashGaps(0 To numDashGaps - 1) As Long
                
                Dim i As Long
                For i = 0 To numDashGaps - 1
                    lineDashGaps(i) = srcStream.ReadLong()
                Next i
                
            End If
            
            'For now, just create a dummy pen with the specified start and end cap types
            ' (the ones PD supports, at least)
            dstPen.SetPenStartCap GetPDLineCapFromPSPLineCap(startCapType)
            dstPen.SetPenEndCap GetPDLineCapFromPSPLineCap(endCapType)
            
            'TODO: handle cap multipliers here
            
            'TODO: handle custom dash gap array here
            
            'All special line styles have been applied!
            okToProceed = psp_Success
        
        Else
            InternalError funcName, "bad line-style chunk size: " & chunkSize, srcWarnings
            okToProceed = psp_Warning
            GoTo EarlyTermination
        End If
    
    'Other blocks are *not* valid
    Else
        InternalError funcName, "unknown blockID: " & blockID, srcWarnings
        okToProceed = psp_Warning
        GoTo EarlyTermination
    End If
    
EarlyTermination:
    BuildBrushForPaintBlock = okToProceed
    
    'Align the stream using the original block header length
    srcStream.SetPosition endOfBlockHeader + blockLength, FILE_BEGIN
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description, srcWarnings
    srcWarnings.AddString "Internal error in pdPSPShape." & funcName & ", #" & Err.Number & ": " & Err.Description
    BuildBrushForPaintBlock = psp_Failure
    
End Function

Private Function BuildBrush_Color(ByRef srcStream As pdStream, ByRef srcWarnings As pdStringStack, ByRef dstBrush As pd2DBrush) As PD_PSPResult

    On Error GoTo InternalVBError
    Const funcName As String = "BuildBrush_Color"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    Dim origStreamPos As Long, chunkSize As Long
    origStreamPos = srcStream.GetPosition()
    chunkSize = srcStream.ReadLong()
    
    If (chunkSize > 0) Then
    
        'All that follows is an RGB color definition, and a color palette index.
        ' Handling a color palette index is TODO
        Dim newColor As RGBQuad, palIndex As Long
        srcStream.ReadBytesToBarePointer VarPtr(newColor), 4
        palIndex = srcStream.ReadLong()
        'If (palIndex >= 0) Then...
        okToProceed = psp_Success
    
    Else
        InternalError funcName, "null color brush chunk size", srcWarnings
        okToProceed = psp_Warning
        GoTo EarlyTermination
    End If
    
EarlyTermination:
    BuildBrush_Color = okToProceed
    
    'Align the stream using the original block header length
    srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description, srcWarnings
    srcWarnings.AddString "Internal error in pdPSPShape." & funcName & ", #" & Err.Number & ": " & Err.Description
    BuildBrush_Color = psp_Failure
    
End Function

Private Function BuildPD2DPolygon(ByRef srcWarnings As pdStringStack, ByRef srcHeader As PSPImageHeader) As PD_PSPResult

    On Error GoTo InternalVBError
    Const funcName As String = "BuildPD2DPolygon"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'Ensure we have actual nodes to work with
    If (m_NodeCount <= 1) Then
        InternalError funcName, "no nodes for polygon constructor", srcWarnings
        okToProceed = psp_Warning
        GoTo EarlyTermination
    End If
    
    'Our next goal is to convert all PSP-format polyline nodes to something a pd2D path object
    ' can work with.  The PSP spec doesn't describe vector contents in any kind of detail,
    ' so this implementation is entirely based on detective work (...or more accurately,
    ' blind guesswork lol)
    Set m_PolylineShape = New pd2DPath
    m_PolylineShape.SetFillRule P2_FR_Winding
    
    'PSP vector data works as a list of nodes, with move-to commands like GDI.  GDI+ works on lines
    ' and sub-paths, with "stop this sub-path / start a new sub-path" commands instead.  We need to
    ' convert between these formats as we go.
    Dim i As Long
    For i = 0 To m_NodeCount - 1
        
        'Skip invisible nodes
        If ((m_Nodes(i).pn_NodeFlags And keNodeVisible) = 0) Then GoTo NextPoint
        
        'When I first implemented this class, it was helpful to display node attributes.
        ' Here's the code in case you need to review parser behavior in the future:
        'PDDebug.LogAction i & ": " & m_Nodes(i).pn_MoveTo & " - " & m_Nodes(i).pn_X & ", " & m_Nodes(i).pn_Y & " - " & m_Nodes(i).pn_hX1 & ", " & m_Nodes(i).pn_hY1 & " - " & m_Nodes(i).pn_hX2 & ", " & m_Nodes(i).pn_hY2
        
        'On the first point, always start a new sub-path
        If (i = 0) Then
            m_PolylineShape.StartNewFigure
        
        'n > 1
        Else
        
            'Check the current command.  If it is a MoveTo command, end the current sub-path
            ' (because we're about to start a new one).
            If m_Nodes(i).pn_MoveTo Then
                m_PolylineShape.CloseCurrentFigure
            
            'This is not a MoveTo command.  Connect this point to the one in front of it.
            Else
                
                'If the previous point was a MoveTo command, start a new sub-figure.  (We handle
                ' MoveTo commands like this to prevent errors if a path were to somehow define
                ' multiple MoveTo commands in a row.)  Note, however, that we don't need to do
                ' this on the first node; it's covered by the default i = 0 case, above.
                If (m_Nodes(i - 1).pn_MoveTo And (i <> 1)) Then m_PolylineShape.StartNewFigure
                
                'When debugging, it may be helpful to just add nodes to the current figure as
                ' straight lines.  Here's the code to do that:
                'm_PolylineShape.AddLine m_Nodes(i - 1).pn_X, m_Nodes(i - 1).pn_Y, m_Nodes(i).pn_X, m_Nodes(i).pn_Y
                
                'Add the current figure as bezier curves.
                '
                'Note that PSP control point definition is a little sketchy.  It took no small amount of
                ' trial-and-error testing to figure out that you need to use one control point from the
                ' previous point, and one control point from the current point to make node descriptors
                ' work like every other graphics library.  (PSPs save node data in a baffling order,
                ' but if you mix control points from neighboring nodes, the end result looks correct!)
                m_PolylineShape.AddBezierCurve m_Nodes(i - 1).pn_X, m_Nodes(i - 1).pn_Y, m_Nodes(i - 1).pn_hX2, m_Nodes(i - 1).pn_hY2, m_Nodes(i).pn_hX1, m_Nodes(i).pn_hY1, m_Nodes(i).pn_X, m_Nodes(i).pn_Y
                
                'If this is the end of the current point collection, close the current path, since there
                ' are no more points to add.
                If (i = m_NodeCount - 1) Then m_PolylineShape.CloseCurrentFigure
                
            End If
            
        End If
        
NextPoint:
    Next i
    
EarlyTermination:
    BuildPD2DPolygon = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description, srcWarnings
    srcWarnings.AddString "Internal error in pdPSPShape." & funcName & ", #" & Err.Number & ": " & Err.Description
    BuildPD2DPolygon = psp_Failure
    
End Function

Private Function GetPDLineCapFromPSPLineCap(ByVal pspCapType As PSPStyleCapType) As PD_2D_LineCap
    
    'Set a default; PSP currently defines many line caps that PD does not.
    ' Caps without a PD analog will get this type of cap.
    GetPDLineCapFromPSPLineCap = P2_LC_Flat
    
    'NOTE: PSP6 has a *different* enum.  Solving that is TODO!
    Select Case pspCapType
        Case keSCTCapFlat
            GetPDLineCapFromPSPLineCap = P2_LC_Flat
        Case keSCTCapRound
            GetPDLineCapFromPSPLineCap = P2_LC_Round
        Case keSCTCapSquare
            GetPDLineCapFromPSPLineCap = P2_LC_Square
        Case keSCTCapArrow
        Case keSCTCapCadArrow
        Case keSCTCapCurvedTipArrow
        Case keSCTCapRingBaseArrow
        Case keSCTCapFluerDelis
        Case keSCTCapFootball
        Case keSCTCapXr71Arrow
        Case keSCTCapLilly
        Case keSCTCapPinapple
        Case keSCTCapBall
            GetPDLineCapFromPSPLineCap = P2_LC_RoundAnchor
        Case keSCTCapTulip
        Case Else
            PDDebug.LogAction "WARNING: pdPSPShape.GetPDLineCapFromPSPLineCap() encountered an unknown line cap: " & pspCapType
    End Select
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, ByRef parentWarningStack As pdStringStack, Optional ByVal writeDebugLog As Boolean = True)
    
    Dim errText As String
    errText = "WARNING! pdPSPShape." & funcName & "() reported an error: " & errDescription
    If (Not parentWarningStack Is Nothing) Then parentWarningStack.AddString errText
    
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction errText
    Else
        Debug.Print errText
    End If
    
End Sub

