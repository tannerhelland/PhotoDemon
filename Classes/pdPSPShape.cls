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
'Copyright 2020-2026 by Tanner Helland
'Created: 31/December/20
'Last updated: 25/May/22
'Last update: in text layers, replace standalone LF chars with CRLF to ensure correct linebreak behavior
'
'This class describes a single vector "shape" inside a JASC/Corel Paint Shop Pro image file.
' It has been custom-built for PhotoDemon, with a special emphasis on parsing performance.
' Vector rendering elements are automatically translated to pd2D as needed (and pd2D ultimately
' handles all rendering).
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

'When debugging vector layers, it can be helpful to simplify vector shapes to confirm that
' everything is working as intended.  These debug modes must *not* be activated in production builds.
Private Const PSP_VECTOR_DEBUG_NO_CURVES As Boolean = False
Private Const PSP_VECTOR_DEBUG_RNDCOLORS As Boolean = False

'This randomize object is also used purely for debugging.  You can remove it in production code.
Private m_Randomize As pdRandomize

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
    keStyleNone = &H0&           '// No paint style info applies
    keStyleColor = &H1&          '// Color paint style info
    keStyleGradient = &H2&       '// Gradient paint style info
    keStylePattern = &H4&        '// Pattern paint style info
    keStylePaper = &H8&          '// Paper paint style info
    keStylePen = &H10&           '// Organic pen paint style info
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
' (describing node count, basically), then a variable-length list of nodes, including control points.
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

'Text layers have their own custom attribute header, followed by a similar definition header
' (describing node count, same as polylines), then a variable-length list of text element nodes
Private Type PSP_TextAttributes
    ta_Size As Long
    ta_Alignment As PSPTextAlignment
    ta_XInsertPt As Long
    ta_YInsertPt As Long
    ta_DeformationMatrix(0 To 8) As Double  '3x3 deformation matrix
    
    'Members past this point may not exist; use SIZE to confirm existence
    
    ta_TextFlow As PSPTextFlow
    ta_PathOffset As Double     'Only used for text-on-path; describes offset from path to text
    'Future expansion fields possible; use Size member to skip
End Type

Private m_TextAttributes As PSP_TextAttributes

'After text attributes, PSPs encode a variable-length list of individual text "elements".
' These can be single-characters, character-style-definitions (fonts, color, etc), or
' line-style-definitions (the spec only defines "leading width" but allows for future expansion).
' PD decodes these to an array of the following struct, but note that various struct members
' are only loaded if they're relevant to a given element type.  Said another way, you *MUST*
' query the type before attempting to use anything from the rest of the struct!
Private Type PSP_TextElement
    te_Type As PSPTextElementType
    
    'Only relevant for character definitions:
    te_CharCode As Long     'The spec calls this "the Unicode value of the character" which is silly - I assume UTF-32 is the actual encoding?
    
    'Only relevant for line definitions:
    te_TextLeadingValue As Long
    
    'Only relevant for character style definitions, and not all definitions are guaranteed to
    ' exist depending on the version of the PSP file (e.g. PSP 6 doesn't describe as many
    ' attributes as PSP 7 does):
    te_FontName As String
    te_CharProperties As PSPCharacterProperties
    te_CharWeight As Long
    te_CharSet As Long      'Not currently supported in PD, and it's meaning is unclear given that the text
                            ' is supposedly Unicode, which is incompatible with the notion of character sets??
    te_FontSize As Long
    te_AntialiasMode As PSPAntialiasMode
    te_Justify As Byte      'Enum isn't described in the spec?
    te_AutoKern As Boolean
    te_KerningValue As Double
    te_TrackingValue As Double
    te_LineLeadingValue As Double
    te_Stroked As Boolean
    te_Filled As Boolean
    te_Styled As Boolean
    te_StrokeWidth As Double
    te_StartCapType As PSPStyleCapType
    te_StartCapMultiplier As Boolean
    te_StartCapWidthMultiplier As Double
    te_StartCapHeightMultiplier As Double
    te_EndCapType As PSPStyleCapType
    te_EndCapMultiplier As Boolean
    te_EndCapWidthMultiplier As Double
    te_EndCapHeightMultiplier As Double
    te_LineJoin As PSPStyleJoinType
    te_MiterLimit As Double
    'Future expansion fields possible; PD doesn't track a Size member for this struct
    ' (because there are actually several different size members throughout the struct)
    ' but it will be accounted for during parsing
End Type

Private m_numTextElements As Long, m_TextElements() As PSP_TextElement

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

'Gradient brushes have a custom header and specialized attributes
Private Type PSP_GradientChunk
    gc_Size As Long
    gc_Color As RGBQuad
    gc_Location As Long
    gc_Midpoint As Long
End Type

Private Type PSP_GradientDefinition
    gd_Size As Long
    gd_Name As String
    gd_ID As Long
    gd_Invert As Boolean
    gd_cX As Long
    gd_cY As Long
    gd_fX As Long
    gd_fY As Long
    gd_Angle As Double  'Linear gradients only
    gd_Repeats As Long
    gd_Type As PSPStyleGradientType
    gd_NumColorStops As Long
    gd_NumOpacityStops As Long
    gd_Colors() As PSP_GradientChunk
    gd_Opacity() As PSP_GradientChunk   'Only the A value is filled, and it's on the range [0, 100]
    
    'PD-specific; this holds a merged RGBA list of both colors *and* opacity
    gd_NumFinalStops As Long
    gd_Final() As GradientPoint
End Type

Private m_GradientDefinition As PSP_GradientDefinition

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

'Files from different PSP versions use totally different structs in places.  Because the PSP spec
' is incredibly ill-conceived, there is no natural way to identify these changes.  Instead, we can
' only rely on the header version of *the entire goddamn file* to try and pick out these files if
' they occur.  (What makes this even more infuriating is their "chunk" file layout explicitly
' allows for expansion slots at the end of each chunk - but instead of using these, they cram new
' data types into the middle of the struct, in structs that already have variable-length members,
' so it's impossible to detect format changes by all the built-in mechanisms they designed
' EXPLICITLY FOR DETECTING FORMAT CHANGES.  Seriously, this makes the PSD spec look like it was
' invented by savants by comparison...)
Private m_FilePSPVersion As Long, m_FileIsPSPv6 As Boolean, m_FileIsPSPv8 As Boolean

'Do *not* call this unless the underlying shape is a text object, obviously!
' If successful, the function returns the ID (not the index) of the newly created layer.
Friend Function CreateTextLayerNow(ByRef dstImage As pdImage, ByRef dstLayer As pdPSPLayer) As Long
    
    Const funcName As String = "CreateTextLayerNow"
    
    CreateTextLayerNow = -1
    
    If (m_numTextElements <= 0) Then
        InternalError funcName, "no text elements to render!", Nothing
        Exit Function
    End If
    
    'Before creating the actual text layer, we need to do some housekeeping on our existing
    ' text element list.
    
    'For one thing, we first need to produce a contiguous string from the individual
    ' character collection!
    Dim finalString As pdString
    Set finalString = New pdString
    
    'We also need to determine things like font name, color, etc.  PSP supports multiple styles
    ' within a single text layer.  PhotoDemon does not.  For now, all we can do is take the
    ' first style in a text layer, and apply that to the layer as a whole.
    Dim styleIndex As Long
    styleIndex = -1
    
    Dim i As Long
    For i = 0 To m_numTextElements - 1
        If (m_TextElements(i).te_Type = keTextElemChar) Then
            finalString.Append ChrW$(m_TextElements(i).te_CharCode)
        ElseIf (m_TextElements(i).te_Type = keTextElemLineStyle) Then
            'not used at present
        ElseIf (m_TextElements(i).te_Type = keTextElemCharStyle) Then
            If (styleIndex < 0) Then styleIndex = i
        Else
            'unknown element block; can't do much with it at present
        End If
    Next i
    
    'With the string assembled, we can actually place it into a text layer
    Dim textContents As String
    textContents = finalString.ToString()
    
    'Look for line-breaks and replace as necessary.  (PSP files appear to encode LF only, and we want
    ' standard Windows CRLF.)
    If (InStr(1, textContents, vbLf, vbBinaryCompare) <> 0) Then
        If (InStr(1, textContents, vbCr, vbBinaryCompare) = 0) Then
            textContents = Replace$(textContents, vbLf, vbCrLf, Compare:=vbBinaryCompare)
        End If
    End If
    
    'Ask the parent pdImage to create a new layer object
    Dim newLayerID As Long
    newLayerID = dstImage.CreateBlankLayer()
    
    'Create a placeholder DIB for the new layer (the parent image will regenerate this based
    ' on the text contents, but it needs a valid DIB to start)
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateBlank 1, 1, 32, 0, 0
    tmpDIB.SetInitialAlphaPremultiplicationState True
    
    'Assign the newly created DIB and layer name to the layer object, and mirror the
    ' blend mode and opacity of the parent layer.  (This doesn't always produce identical
    ' results due to the way PSP groups layers, but it's a good approximation on most images.)
    dstImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_TextAdvanced, textContents, tmpDIB
    dstImage.GetLayerByID(newLayerID).SetLayerBlendMode dstLayer.GetLayerBlendMode()
    dstImage.GetLayerByID(newLayerID).SetLayerOpacity dstLayer.GetLayerOpacity()
    dstImage.GetLayerByID(newLayerID).SetLayerVisibility ((m_ShapeAttributes.sa_Flags And keShapeVisible) <> 0)
    
    'Apply the text itself
    dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_Text, textContents
    
    'Apply other text settings from the header
    Dim convertedTextAlignment As GP_StringAlignment
    Select Case m_TextAttributes.ta_Alignment
        Case keTextAlignmentLeft
            convertedTextAlignment = StringAlignmentNear
        Case keTextAlignmentCenter
            convertedTextAlignment = StringAlignmentCenter
        Case keTextAlignmentRight
            convertedTextAlignment = StringAlignmentFar
        
        'Failsafe for future PSP changes; assume justified text, I guess?
        Case Else
            convertedTextAlignment = StringAlignmentJustify
    End Select
    
    dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_HorizontalAlignment, convertedTextAlignment
    dstImage.GetLayerByID(newLayerID).NotifyOfDestructiveChanges
    
    'Layer width/height requires font data to solve.  It may not be 100% accurate, owing to rendering
    ' differences between PD and PSP.
    Dim tmpFont As pdFont
    Set tmpFont = New pdFont
    tmpFont.SetFontFace m_TextElements(styleIndex).te_FontName
    
    'Other font parameters also need to be set to calculate layout semi-correctly
    tmpFont.SetFontBold (m_TextElements(styleIndex).te_CharWeight > 400)    '400 = normal, 700 = bold
    tmpFont.SetFontItalic ((m_TextElements(styleIndex).te_CharProperties And keStyleItalic) <> 0)
    tmpFont.SetFontUnderline ((m_TextElements(styleIndex).te_CharProperties And keStyleUnderlined) <> 0)
    tmpFont.SetFontStrikeout ((m_TextElements(styleIndex).te_CharProperties And keStyleStruck) <> 0)
    
    'Font size calculation is complicated because PSP lets you arbitrarily stretch a text layer
    ' in either direction, so font size is basically meaningless.  We attempt to approximate the
    ' desired font size by multiplying the embedded font size by the average width/height multiplier
    ' of the deformation matrix.  (If the user did *not* resize the PSP layer at all, this will
    ' produce roughly identical size results.)
    Dim fontSizeInPixels As Single
    fontSizeInPixels = m_TextElements(styleIndex).te_FontSize * ((m_TextAttributes.ta_DeformationMatrix(0) + m_TextAttributes.ta_DeformationMatrix(4)) / 2)
    tmpFont.SetFontSize fontSizeInPixels * 0.75     '(72 / 96 DPI)
    
    'PSP doesn't actually store text layer offsets.  Instead, it stores a 3x3 deformation matrix
    ' from which we can roughly reverse-engineer the expected baseline position of the first
    ' line of text.  (Note that the baseline position only considers the font *ascent*, not its
    ' *descent*, and I have no idea what font library PSP uses so these results are *not* identical.)
    Dim srcFontMetrics As TEXTMETRIC
    tmpFont.AttachToDC tmpDIB.GetDIBDC()
    Fonts.FillTextMetrics tmpDIB.GetDIBDC, srcFontMetrics
    
    'You'll notice a few "+2" additions on font boundaries in this function.  Modern hinting
    ' algorithms produce rounder characters with improved outline accuracy, which can lead to
    ' small differences in boundary sizes vs PSP itself.  To ensure boundaries are (roughly)
    ' similar to PSP, we expand font boundary sizes by a pixel in either direction to account
    ' for this.
    Dim approxFontWidth As Single
    approxFontWidth = tmpFont.GetWidthOfString(textContents) + 2!
    
    Dim boundaryRect As RectL
    tmpFont.GetBoundaryRectOfMultilineString textContents, approxFontWidth, boundaryRect, True
    tmpFont.ReleaseFromDC
    
    dstImage.GetLayerByID(newLayerID).SetLayerWidth (boundaryRect.Right - boundaryRect.Left) + 2
    dstImage.GetLayerByID(newLayerID).SetLayerHeight (boundaryRect.Bottom - boundaryRect.Top) + 2
    
    'Apply layer (x, y) offsets
    dstImage.GetLayerByID(newLayerID).SetLayerOffsetX m_TextAttributes.ta_DeformationMatrix(2)
    dstImage.GetLayerByID(newLayerID).SetLayerOffsetY (m_TextAttributes.ta_DeformationMatrix(5) - srcFontMetrics.tmAscent)   'boundaryRect.Bottom    ' approxFontHeight  'Note: offset is to BASELINE; gotta fix this
    
    'Trial-and-error leads me to believe that in PSP v6 files, center-aligned text now positions along
    ' the text-box's center line.  (TBD - determining if something similar happens for right-aligned text.)
    If m_FileIsPSPv8 Then
        If (convertedTextAlignment = StringAlignmentCenter) Then dstImage.GetLayerByID(newLayerID).SetLayerOffsetX m_TextAttributes.ta_DeformationMatrix(2) - ((boundaryRect.Right - boundaryRect.Left) + 2) \ 2
    End If
    
    'Apply other font properties
    dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_FontFace, tmpFont.GetFontFace()
    dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_FontSize, fontSizeInPixels
    dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_FontSizeUnit, fu_Pixel
    dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_FontBold, tmpFont.GetFontBold()
    dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_FontItalic, tmpFont.GetFontItalic()
    dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_FontUnderline, tmpFont.GetFontUnderline()
    dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_FontStrikeout, tmpFont.GetFontStrikeout()
    
    'Set fill/stroke brushes
    dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_FillActive, m_TextElements(styleIndex).te_Filled
    If m_TextElements(styleIndex).te_Filled Then dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_FillBrush, m_BrushFill.GetBrushPropertiesAsXML()
    
    'Text stroking is TODO (we need to create a pen from the PSP brush)
    'If m_TextElements(styleIndex).te_Stroked Then dstImage.GetLayerByID(newLayerID).SetTextLayerProperty ptp_OutlinePen, m_BrushFill
    
    CreateTextLayerNow = newLayerID
    
End Function

Friend Function IsShapeOK() As Boolean
    IsShapeOK = m_ShapeOK
End Function

Friend Function IsShapeText() As Boolean
    IsShapeText = (m_numTextElements > 0)
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
    ' rely on the original version at the head of the file to tell us how to parse.
    m_FilePSPVersion = srcHeader.psph_VersionMajor
    m_FileIsPSPv6 = (srcHeader.psph_VersionMajor <= 4)
    m_FileIsPSPv8 = (srcHeader.psph_VersionMajor >= 6)
    
    'Immediately following the block header is a "shape attributes" chunk.  This tells us
    ' the shape name and type, which lead to different blueprints for parsing the remainder
    ' of this shape block.
    With m_ShapeAttributes
        .sa_Size = srcStream.ReadLong()
        If (.sa_Size < 4) Then
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
    
        'Text layers in PSP are treated as a subset of "vector shapes"
        Case keVSTText
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Found shape: text"
            okToProceed = ReadVectorText(srcStream, srcWarnings, srcHeader)
            m_ShapeOK = (okToProceed = psp_Success)
            
        'These all use polyline definitions "under the hood"
        Case keVSTPolyline, keVSTEllipse, keVSTPolygon
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Found shape: polyline"
            okToProceed = ReadPolylines(srcStream, srcWarnings, srcHeader)
            m_ShapeOK = (okToProceed = psp_Success)
            
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
            If (sLen > 0) Then m_ShapeAttributes.sa_Name = srcStream.ReadString_UTF8(sLen)
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
            If (m_FilePSPVersion = 5) Then
                .pa_StrokeWidth = srcStream.ReadIntUnsigned()
            Else
                .pa_StrokeWidth = srcStream.ReadDouble()
            End If
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
    ' way PSP embeds chunk vs block sizes (this function parses a chunk, not a block).
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
    
    'If a polygon wasn't produced (this will be true for groups that consist of only text layers,
    ' for example), skip this render step
    If (m_PolylineShape Is Nothing) Then
        RenderShapeIntoDIB = False
        Exit Function
    End If
    
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
     
    'PDDebug.LogAction "rendering shape: " & m_ShapeAttributes.sa_Name & ", " & m_PolylineShape.GetNumPathPoints()
    
    Dim polylineBoundaries As RectF
    polylineBoundaries = m_PolylineShape.GetPathBoundariesF()
    
    'Debug mode is helpful for studying vector data as read from file, without the complications
    ' of brush and pen parsing+generation.
    If PSP_VECTOR_DEBUG_RNDCOLORS Then
        
        Dim tmpBrush As pd2DBrush, tmpPen As pd2DPen
        
        Dim randomColor As Long
        randomColor = m_Randomize.GetRandomInt_VB()
        
        Drawing2D.QuickCreateSolidBrush tmpBrush, randomColor, 100!
        PD2D.FillPath dstSurface, tmpBrush, m_PolylineShape
        
        Drawing2D.QuickCreateSolidPen tmpPen, 1!, (Not randomColor) And &HFFFFFF, 100!, P2_LJ_Round
        PD2D.DrawPath dstSurface, tmpPen, m_PolylineShape
        
        RenderShapeIntoDIB = True
        Exit Function
        
    End If
    
    'Fill if flagged
    If m_PolylineAttributes.pa_Filled Then
        
        'PDDebug.LogAction "filling..." & m_BrushFill.GetBrushOpacity
        
        'Make sure our fill brush exists
        If (Not m_BrushFill Is Nothing) Then
            If (m_BrushFill.GetBrushMode = P2_BM_Gradient) Then m_BrushFill.SetBoundaryRect m_PolylineShape.GetPathBoundariesF()
            
            PD2D.FillPath dstSurface, m_BrushFill, m_PolylineShape
            RenderShapeIntoDIB = True
        Else
            InternalError funcName, "no fill brush", Nothing
        End If
        
    End If
    
    'Stroke if flagged
    If m_PolylineAttributes.pa_Stroked Then
        
        'PDDebug.LogAction "stroking..."
        
        If (Not m_BrushStroke Is Nothing) Then
        
            Dim shapePen As pd2DPen
            Set shapePen = New pd2DPen
            
            'Apply separate pen settings from the *parent* shape object
            shapePen.SetPenWidth m_PolylineAttributes.pa_StrokeWidth
            
            'Create a pen from the stroking brush specified in the file.
            ' (Note that gradient pens are TODO; they will require us to calculate a boundary somehow)
            If (m_BrushStroke.GetBrushMode = P2_BM_Solid) Then
                shapePen.SetPenColorRGBA m_BrushStroke.GetBrushColorRGBA
                shapePen.SetPenOpacity 100!
            Else
                If (m_BrushFill.GetBrushMode = P2_BM_Gradient) Then m_BrushFill.SetBoundaryRect m_PolylineShape.GetPathBoundariesF(shapePen)
                shapePen.CreatePenFromBrush m_BrushStroke
            End If
            
            'If a special "styled" pen exists, use it as our base pen.
            ' NOTE: this needs to be revisited; it won't work because we're creating a pen from
            ' a brush, so order needs to be observed
            'If (Not m_PenStyled Is Nothing) Then
            '    PDDebug.LogAction "cloning style pen"
            '    shapePen.ClonePen m_PenStyled
            'Else
                shapePen.SetPenStartCap GetPDLineCapFromPSPLineCap(m_PolylineAttributes.pa_StartCapType)
                shapePen.SetPenEndCap GetPDLineCapFromPSPLineCap(m_PolylineAttributes.pa_EndCapType)
                'Cap multipliers are TODO?
            'End If
            
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
            
            'Finally, perform the damn stroke!
            PD2D.DrawPath dstSurface, shapePen, m_PolylineShape
            
            RenderShapeIntoDIB = True
        
        Else
            InternalError funcName, "no stroke brush", Nothing
        End If
        
    End If
    
End Function

'Text layers bear strong similarity to polyline layers, and PSP treats text as a subset of "vector shapes"
Private Function ReadVectorText(ByRef srcStream As pdStream, ByRef srcWarnings As pdStringStack, ByRef srcHeader As PSPImageHeader) As PD_PSPResult
    
    On Error GoTo InternalVBError
    Const funcName As String = "ReadVectorText"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    'Stream position at entry.  Note that this is a *chunk*, not a *block* so we can't use this
    ' for absolute alignment.  (It is, however, helpful for debug purposes.)
    Dim origStreamPos As Long
    origStreamPos = srcStream.GetPosition()
    
    Dim chunkSize As Long
    chunkSize = srcStream.ReadLong()
    If (chunkSize < 4) Then
        InternalError funcName, "bad text attribute size: " & chunkSize & ", " & srcStream.GetPosition(), srcWarnings
        okToProceed = psp_Failure
        GoTo EarlyTermination
    End If
    
    'This step starts out pretty simple - just fill the attribute header!
    With m_TextAttributes
        
        .ta_Alignment = srcStream.ReadByte()
        .ta_XInsertPt = srcStream.ReadLong()
        .ta_YInsertPt = srcStream.ReadLong()
        srcStream.ReadBytesToBarePointer VarPtr(.ta_DeformationMatrix(0)), 9 * 8 '9 double-width floats
        
        'Subsequent values may not exist; header is inconsistent between versions,
        ' so use chunk size as a guide
        If (chunkSize >= 86) Then .ta_TextFlow = srcStream.ReadByte()
        If (chunkSize >= 94) Then .ta_PathOffset = srcStream.ReadDouble()
        
    End With
    
    'Align stream using chunk size
    srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
    
    'Next comes text shape definition.  In public specs, this simply defines the number
    ' of text "element" chunks that follow.
    origStreamPos = srcStream.GetPosition()
    chunkSize = srcStream.ReadLong()
    If (chunkSize < 8) Then
        InternalError funcName, "bad text definition size: " & chunkSize, srcWarnings
        okToProceed = psp_Failure
        GoTo EarlyTermination
    End If
    
    m_numTextElements = srcStream.ReadLong()
    If (m_numTextElements <= 0) Then
        InternalError funcName, "bad text element count: " & m_numTextElements, srcWarnings
        okToProceed = psp_Failure
        GoTo EarlyTermination
    Else
        If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "found " & m_numTextElements & " text elements (" & srcStream.GetPosition() & ", " & chunkSize & ")"
    End If
    ReDim m_TextElements(0 To m_numTextElements - 1) As PSP_TextElement
    
    'Again, align the stream using chunk size
    srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
    
    'We now need to read in (n) text element chunks.  This count is deceptive because there
    ' are actually *TWO* chunks per text element - a quasi-header "attributes" chunk, which
    ' tells us which kind of chunk follows (a Unicode character or a character/line style
    ' description, which is important because PSP files support multi-style text layers),
    ' then a "definition" chunk whose layout depends on the value of the preceding
    ' "attributes" chunk.
    
    'Again, because the PSP spec is very badly designed, these are all just separated
    ' by variable-length chunk length descriptors - no IDs anywhere - so you just gotta
    ' pray that the file encoder did everything correctly, because any sort of validation
    ' of stream alignment (or content correctness) is effectively impossible.
    Dim i As Long
    For i = 0 To m_numTextElements - 1
        
        'Parse the "header" text element attributes chunk
        origStreamPos = srcStream.GetPosition()
        chunkSize = srcStream.ReadLong()
        If (chunkSize >= 6) Then
            m_TextElements(i).te_Type = srcStream.ReadIntUnsigned()
        Else
            InternalError funcName, "bad text element attribute size: " & chunkSize & ", " & srcStream.GetPosition(), srcWarnings
            okToProceed = psp_Failure
            GoTo EarlyTermination
        End If
        
        'Realign the pointer using chunk size
        srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
        
        'Retrieve new chunk size
        origStreamPos = srcStream.GetPosition()
        chunkSize = srcStream.ReadLong()
        If (chunkSize < 4) Then
            InternalError funcName, "bad text element definition size: " & chunkSize, srcWarnings
            okToProceed = psp_Failure
            GoTo EarlyTermination
        Else
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Element definition size (" & i & "): " & chunkSize
        End If
        
        'Split further handling based on text element type
        Select Case m_TextElements(i).te_Type
            
            'Characters are easy
            Case keTextElemChar
                m_TextElements(i).te_CharCode = srcStream.ReadLong()
                
                'Always align against the original chunk size
                srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
                
            'Character styles are *not* easy
            Case keTextElemCharStyle
                
                Dim lenFontName As Long
                lenFontName = srcStream.ReadIntUnsigned()
                With m_TextElements(i)
                    
                    'Control for vertical-orientation fonts (which PD deliberately hides from the UI)
                    .te_FontName = srcStream.ReadString_UTF8(lenFontName)
                    If (Left$(.te_FontName, 1) = "@") Then .te_FontName = Right$(.te_FontName, Len(.te_FontName) - 1)
                    
                    .te_CharProperties = srcStream.ReadLong()
                    .te_CharWeight = srcStream.ReadLong()
                    .te_CharSet = srcStream.ReadLong()
                    .te_FontSize = srcStream.ReadLong()
                    
                    'Handling beyond this point changes by PSP file version
                    If (m_FilePSPVersion = 4) Then
                        .te_AutoKern = (srcStream.ReadByte() <> 0)
                        .te_KerningValue = srcStream.ReadDouble()
                        
                        'Manually populate filled/stroked flags because the file does not define these
                        .te_Stroked = False
                        .te_Filled = True
                        .te_Styled = False
                    
                    Else
                    
                        If (m_FilePSPVersion = 5) Then
                            .te_AutoKern = (srcStream.ReadByte() <> 0)
                            .te_KerningValue = srcStream.ReadDouble()
                        
                        'It's entirely possible that modern PSP versions deviate from the last-available
                        ' version of the PSP spec (PSPv8).
                        Else
                            .te_AntialiasMode = srcStream.ReadByte()
                            .te_Justify = srcStream.ReadByte()
                            .te_AutoKern = (srcStream.ReadByte() <> 0)
                            .te_KerningValue = srcStream.ReadDouble()
                            .te_TrackingValue = srcStream.ReadDouble()
                            .te_TextLeadingValue = srcStream.ReadDouble()
                        End If
                        
                        'The remaining values appear to be shared between all versions 5+
                        .te_Stroked = (srcStream.ReadByte() <> 0)
                        .te_Filled = (srcStream.ReadByte() <> 0)
                        .te_Styled = (srcStream.ReadByte() <> 0)
                        .te_StrokeWidth = srcStream.ReadIntUnsigned()
                        .te_StartCapType = srcStream.ReadByte()
                        .te_StartCapMultiplier = (srcStream.ReadByte() <> 0)
                        .te_StartCapWidthMultiplier = srcStream.ReadDouble()
                        .te_StartCapHeightMultiplier = srcStream.ReadDouble()
                        .te_EndCapType = srcStream.ReadByte()
                        .te_EndCapMultiplier = (srcStream.ReadByte() <> 0)
                        .te_EndCapWidthMultiplier = srcStream.ReadDouble()
                        .te_EndCapHeightMultiplier = srcStream.ReadDouble()
                        .te_LineJoin = srcStream.ReadByte()
                        .te_MiterLimit = srcStream.ReadDouble()
                        
                    End If
                    
                End With
                
                'Always align against the original chunk size
                srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
                
                'A character style block is immediately followed by one or more paint blocks.
                
                'v6 only supports a fill block
                If m_FileIsPSPv6 Then
                    If (okToProceed < psp_Failure) Then okToProceed = BuildBrushForPaintBlock(srcStream, srcWarnings, m_BrushFill, Nothing)
                
                'Post-v6 support all three style blocks
                Else
                    If (okToProceed < psp_Failure) Then okToProceed = BuildBrushForPaintBlock(srcStream, srcWarnings, m_BrushStroke, Nothing)
                    If (okToProceed < psp_Failure) Then okToProceed = BuildBrushForPaintBlock(srcStream, srcWarnings, m_BrushFill, Nothing)
                    If (okToProceed < psp_Failure) Then okToProceed = BuildBrushForPaintBlock(srcStream, srcWarnings, Nothing, m_PenStyled)
                End If
                
            'Line styles are easy
            Case keTextElemLineStyle
                m_TextElements(i).te_LineLeadingValue = srcStream.ReadLong()
                
                'Always align against the original chunk size
                srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
            
            'keTextElemUnknown is a member of the enum, but I don't know what to do with it...
            Case Else
                If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "unknown text element encountered (#" & i & "): " & m_TextElements(i).te_Type
                
                'Always align against the original chunk size
                srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
                
        End Select
        
    Next i
    
    'Unlike most other read functions in PD's PSP-centric collection, this reader does
    ' *not* align the stream before exiting.  That is up to our parent class due to the
    ' way PSP embeds chunk vs block sizes (this function parses a chunk, not a block).
EarlyTermination:
    ReadVectorText = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description, srcWarnings
    srcWarnings.AddString "Internal error in pdPSPShape." & funcName & ", #" & Err.Number & ": " & Err.Description
    ReadVectorText = psp_Failure
    
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
        
        'DEBUG ONLY: prep an opaque, black brush as a fallback in case other methods fail
        dstBrush.SetBrushColor vbBlack
        dstBrush.SetBrushOpacity 100!
        
        chunkSize = srcStream.ReadLong()
        If (chunkSize > 0) Then
        
            Dim paintStyleType As PSPPaintStyleType
            paintStyleType = srcStream.ReadIntUnsigned()
            
            'Before handling paint style type, jump to the end of this chunk.
            ' (The spec allows future expansion bytes here, and we must account for them.)
            srcStream.SetPosition endOfBlockHeader + chunkSize, FILE_BEGIN
            
            If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Retrieving paint styles (type: " & Hex$(paintStyleType) & ", position: " & srcStream.GetPosition() & ")"
            
            'Based on paint style type, we'll branch into a new parser.  Note that multiple
            ' brushes can be defined!  This allows e.g. a pattern to be mixed with a gradient.
            If (paintStyleType = keStyleNone) Then
                If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "No paint style defined for this layer (" & chunkSize & "); layer will be ignored"
                BuildBrushForPaintBlock = psp_Warning
            Else
                
                'Color style blocks appear to work across most PSP file versions
                If ((paintStyleType And keStyleColor) <> 0) Then
                    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "solid brush found; retrieving now..."
                    BuildBrushForPaintBlock = BuildBrush_Color(srcStream, srcWarnings, dstBrush)
                End If
                
                'The order of style blocks appears to vary in later PSP versions.  (That, or the format
                ' has changed from what's defined in the spec.)
                If (m_FilePSPVersion <= 6) Then
                    
                    If ((paintStyleType And keStyleGradient) <> 0) Then
                        If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "gradient brush found; retrieving now..."
                        BuildBrushForPaintBlock = BuildBrush_Gradient(srcStream, srcWarnings, dstBrush)
                    End If
                    
                    If ((paintStyleType And keStylePattern) <> 0) Then
                        chunkSize = srcStream.ReadLong()
                        InternalError funcName, "pattern styles are TODO (" & chunkSize & ")", srcWarnings
                        BuildBrushForPaintBlock = psp_Warning
                        If (chunkSize > 4) Then srcStream.SetPosition chunkSize - 4, FILE_CURRENT
                    End If
                    
                    If ((paintStyleType And keStylePaper) <> 0) Then
                        chunkSize = srcStream.ReadLong()
                        InternalError funcName, "paper styles are TODO (" & chunkSize & ")", srcWarnings
                        BuildBrushForPaintBlock = psp_Warning
                        If (chunkSize > 4) Then srcStream.SetPosition chunkSize - 4, FILE_CURRENT
                    End If
                    
                    If ((paintStyleType And keStylePen) <> 0) Then
                        chunkSize = srcStream.ReadLong()
                        InternalError funcName, "organic pen styles are TODO (" & chunkSize & ")", srcWarnings
                        BuildBrushForPaintBlock = psp_Warning
                        If (chunkSize > 4) Then srcStream.SetPosition chunkSize - 4, FILE_CURRENT
                    End If
                
                'At some point, the PSP spec was changed and paint styles are now handled completely differently.
                Else
                    
                    If PSP_DEBUG_VERBOSE Then
                    
                        PDDebug.LogAction "Paint style chunks ----------"
                        Do While (srcStream.GetPosition() < endOfBlockHeader + blockLength)
                            chunkSize = srcStream.ReadLong
                            PDDebug.LogAction chunkSize
                            If (chunkSize > 4) Then srcStream.SetPosition chunkSize - 4, FILE_CURRENT
                        Loop
                        PDDebug.LogAction "End paint style chunks ----------"
                        
                    End If
                    
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
    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Color brush chunk size: " & chunkSize
    
    If (chunkSize > 4) Then
        
        'All that follows is an RGB color definition, and a color palette index.
        ' Handling a color palette index is TODO
        Dim newColor As RGBQuad, palIndex As Long
        newColor.Red = srcStream.ReadByte
        newColor.Green = srcStream.ReadByte
        newColor.Blue = srcStream.ReadByte
        srcStream.ReadByte
        'srcStream.ReadBytesToBarePointer VarPtr(newColor), 4
        newColor.Alpha = 255    'PSP colors do not define opacity (it's handled elsewhere)
        palIndex = srcStream.ReadLong()
        'If (palIndex >= 0) Then...
        okToProceed = psp_Success
        
        'Create a matching pd2D brush
        If (dstBrush Is Nothing) Then Set dstBrush = New pd2DBrush
        dstBrush.SetBrushMode P2_BM_Solid
        dstBrush.SetBrushColorRGBA newColor
        
    Else
        InternalError funcName, "null color brush chunk size", srcWarnings
        okToProceed = psp_Warning
        GoTo EarlyTermination
    End If
    
EarlyTermination:
    BuildBrush_Color = okToProceed
    
    'Align the stream using the original block header length
    If (chunkSize > 4) Then srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description, srcWarnings
    srcWarnings.AddString "Internal error in pdPSPShape." & funcName & ", #" & Err.Number & ": " & Err.Description
    BuildBrush_Color = psp_Failure
    
End Function

Private Function BuildBrush_Gradient(ByRef srcStream As pdStream, ByRef srcWarnings As pdStringStack, ByRef dstBrush As pd2DBrush) As PD_PSPResult

    On Error GoTo InternalVBError
    Const funcName As String = "BuildBrush_Gradient"
    
    Dim okToProceed As PD_PSPResult
    okToProceed = psp_Success
    
    Dim origStreamPos As Long, chunkSize As Long
    origStreamPos = srcStream.GetPosition()
    chunkSize = srcStream.ReadLong()
    If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Gradient brush chunk size: " & chunkSize
    
    If (chunkSize > 4) Then
        
        'Failsafe check for malformed gradient headers or future spec changes.
        ' NOTE: this check is triggered on all PSPs from X6 (chunk size is always, inexplicably, 16).
        ' I don't know how to interpret gradient commands on such files, and I don't know when the
        ' spec changed.  As usual, it's incredibly frustrating and idiotic to even have a spec if
        ' the company using it constantly makes backward-incompatible changes.
        '
        'Regardless, if the chunk is too small to define a gradient (as outlined in the spec),
        ' we'll simply abandon further processing.
        If (chunkSize < 35) Then
            InternalError funcName, "gradient header is too small: " & chunkSize & " (" & origStreamPos & ")", srcWarnings
            srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
            BuildBrush_Gradient = psp_Warning
            Exit Function
        End If
        
        'Read in the gradient definition struct
        With m_GradientDefinition
            .gd_Size = chunkSize
            Dim lenName As Long
            lenName = srcStream.ReadIntUnsigned()
            If (lenName > 0) Then .gd_Name = srcStream.ReadString_UTF8(lenName)
            'If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "Gradient attributes size: " & .gd_Size & ", " & .gd_Name
            .gd_ID = srcStream.ReadLong()
            .gd_Invert = (srcStream.ReadByte() <> 0)
            .gd_cX = srcStream.ReadLong()
            .gd_cY = srcStream.ReadLong()
            If m_FileIsPSPv8 Then
                .gd_fX = srcStream.ReadLong()
                .gd_fY = srcStream.ReadLong()
            End If
            'If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "center/focal:" & .gd_cX & ", " & .gd_cY & ", " & .gd_fX & ", " & .gd_fY
            
            'PSPs use the vertical axis as their angle baseline; rotate the angle accordingly
            .gd_Angle = srcStream.ReadDouble() - 90#
            If (.gd_Angle < 0#) Then .gd_Angle = .gd_Angle + 360#
            
            'Repeats are currently unsupported - TODO
            .gd_Repeats = srcStream.ReadIntUnsigned()
            .gd_Type = srcStream.ReadIntUnsigned()
            'If PSP_DEBUG_VERBOSE Then PDDebug.LogAction "type: " & .gd_Type
            
            'Stops should never be 0, but because an allocation is involved, caution is warranted
            .gd_NumColorStops = srcStream.ReadIntUnsigned()
            If (.gd_NumColorStops > 0) Then
                ReDim .gd_Colors(0 To .gd_NumColorStops - 1) As PSP_GradientChunk
            Else
                InternalError funcName, "no color stops??", srcWarnings
                ReDim .gd_Colors(0) As PSP_GradientChunk
            End If
            
            'Opacity is a separate array of SHORTS on the range [0, 100] (I know, it makes no sense)
            .gd_NumOpacityStops = srcStream.ReadIntUnsigned()
            If (.gd_NumOpacityStops > 0) Then
                ReDim .gd_Opacity(0 To .gd_NumOpacityStops - 1) As PSP_GradientChunk
            Else
                InternalError funcName, "no opacity stops??", srcWarnings
                ReDim .gd_Opacity(0) As PSP_GradientChunk
            End If
            
        End With
        
        'Align the stream
        srcStream.SetPosition origStreamPos + chunkSize, FILE_BEGIN
        
        'Next, a list of gradient color chunks follow.  Load them all.
        Dim i As Long
        If (m_GradientDefinition.gd_NumColorStops > 0) Then
        
            For i = 0 To m_GradientDefinition.gd_NumColorStops - 1
                
                origStreamPos = srcStream.GetPosition
                
                With m_GradientDefinition.gd_Colors(i)
                    .gc_Size = srcStream.ReadLong()
                    .gc_Color.Red = srcStream.ReadByte
                    .gc_Color.Green = srcStream.ReadByte
                    .gc_Color.Blue = srcStream.ReadByte
                    
                    'Alpha is not defined here, but a byte is still allocated for it
                    .gc_Color.Alpha = 255
                    srcStream.ReadByte
                    
                    .gc_Location = srcStream.ReadIntUnsigned()
                    .gc_Midpoint = srcStream.ReadIntUnsigned()
                End With
                
                'Align the stream
                srcStream.SetPosition origStreamPos + m_GradientDefinition.gd_Colors(i).gc_Size
                
            Next i
            
        Else
            InternalError funcName, "no gradient nodes; brush can't be created", srcWarnings
            BuildBrush_Gradient = psp_Warning
            GoTo EarlyTermination
        End If
        
        'Next, a list of opacity chunks follow.  Load them all.
        If (m_GradientDefinition.gd_NumOpacityStops > 0) Then
            
            For i = 0 To m_GradientDefinition.gd_NumOpacityStops - 1
                
                origStreamPos = srcStream.GetPosition
                
                With m_GradientDefinition.gd_Opacity(i)
                    .gc_Size = srcStream.ReadLong()
                    If (m_FileIsPSPv8 Or (.gc_Size > 9)) Then
                        .gc_Color.Alpha = (srcStream.ReadIntUnsigned() And &HFF&)
                    
                    'NOTE: the PSP spec is incorrect here.  On pre-v8 files, opacity stops are
                    ' always 1-byte long.  (Post-v8 files may use a different format entirely,
                    ' potentially stored in the table segment of the file?  Cracking this is
                    ' TODO.)
                    Else
                        .gc_Color.Alpha = srcStream.ReadByte()
                    End If
                    
                    .gc_Location = srcStream.ReadIntUnsigned()
                    .gc_Midpoint = srcStream.ReadIntUnsigned()
                    
                End With
                
                'Align the stream
                srcStream.SetPosition origStreamPos + m_GradientDefinition.gd_Opacity(i).gc_Size
                
            Next i
            
        End If
        
        'With the gradient loaded, we now need to produce a useable brush from it.  First, we need
        ' to merge the separate color and opacity lists into a single RGBA list.
        
        'Start by producing a PD-compatible array of gradient objects from *just* the color values
        ' in the PSP gradient.  (Note that we make this list 2x longer than we need, by default.
        ' This is because PSP gradients support a special "midpoint" value, which can move the
        ' midpoint of the gradient between two nodes.)  We'll trim the list down to the correct
        ' size after we construct it.
        Dim tmpGradientList() As GradientPoint
        ReDim tmpGradientList(0 To (m_GradientDefinition.gd_NumColorStops + 1) * 2) As GradientPoint
        
        Dim curPointIndex As Long, numColorStops As Long
        curPointIndex = 0
        
        Dim tmpR As Long, tmpG As Long, tmpB As Long, tmpA As Long, tmpPos As Single
        
        For i = 0 To m_GradientDefinition.gd_NumColorStops - 1
            
            'Failsafe allocation check
            If (curPointIndex > UBound(tmpGradientList)) Then ReDim Preserve tmpGradientList(0 To curPointIndex * 2) As GradientPoint
            
            'Give the first point a special check.  If its position is *not* 0, we want to create
            ' a 0-position item.
            If (i = 0) And (m_GradientDefinition.gd_Colors(0).gc_Location <> 0) Then
                With m_GradientDefinition.gd_Colors(i)
                    tmpGradientList(curPointIndex).PointRGB = RGB(.gc_Color.Red, .gc_Color.Green, .gc_Color.Blue)
                    tmpGradientList(curPointIndex).PointPosition = 0!
                    tmpGradientList(curPointIndex).PointOpacity = 100!
                End With
                curPointIndex = curPointIndex + 1
            End If
            
            'Add the current point to the collection
            With m_GradientDefinition.gd_Colors(i)
                tmpGradientList(curPointIndex).PointRGB = RGB(.gc_Color.Red, .gc_Color.Green, .gc_Color.Blue)
                tmpGradientList(curPointIndex).PointPosition = .gc_Location / 100!
            End With
            
            'PSP files give opacity its own gradient, separate from the RGB one.  We'll deal with it
            ' in a bit - but for now, default to opaque.
            tmpGradientList(curPointIndex).PointOpacity = 100!
            
            curPointIndex = curPointIndex + 1
            
            'Check the midpoint.  If it is *not* 50, we need to manually add another point to the
            ' collection that splits the distance between this point and the next point in line.
            If (i < m_GradientDefinition.gd_NumColorStops - 1) Then
                If (m_GradientDefinition.gd_Colors(i).gc_Midpoint <> 50) Then
                    If (curPointIndex > UBound(tmpGradientList)) Then ReDim Preserve tmpGradientList(0 To curPointIndex * 2) As GradientPoint
                    
                    With m_GradientDefinition.gd_Colors(i)
                        tmpR = .gc_Color.Red
                        tmpR = tmpR + Int(m_GradientDefinition.gd_Colors(i + 1).gc_Color.Red)
                        tmpG = .gc_Color.Green
                        tmpG = tmpG + Int(m_GradientDefinition.gd_Colors(i + 1).gc_Color.Green)
                        tmpB = .gc_Color.Blue
                        tmpB = tmpB + Int(m_GradientDefinition.gd_Colors(i + 1).gc_Color.Blue)
                        
                        tmpGradientList(curPointIndex).PointRGB = RGB(Int(tmpR / 2 + 0.5), Int(tmpG / 2 + 0.5), Int(tmpB / 2 + 0.5))
                        tmpPos = CSng(m_GradientDefinition.gd_Colors(i + 1).gc_Location - .gc_Location) / 100!
                        tmpGradientList(curPointIndex).PointPosition = (.gc_Location / 100!) + (tmpPos * (.gc_Midpoint / 100!))
                        curPointIndex = curPointIndex + 1
                    End With
                End If
            End If
            
            'If this is the final point and its position isn't 1.0, create a dummy point in its place.
            If (i = m_GradientDefinition.gd_NumColorStops - 1) And (m_GradientDefinition.gd_Colors(i).gc_Location <> 100) Then
                If (curPointIndex > UBound(tmpGradientList)) Then ReDim Preserve tmpGradientList(0 To curPointIndex * 2) As GradientPoint
                With m_GradientDefinition.gd_Colors(i)
                    tmpGradientList(curPointIndex).PointRGB = RGB(.gc_Color.Red, .gc_Color.Green, .gc_Color.Blue)
                    tmpGradientList(curPointIndex).PointPosition = 1!
                    tmpGradientList(curPointIndex).PointOpacity = 100!
                End With
                curPointIndex = curPointIndex + 1
            End If
            
        Next i
        
        numColorStops = curPointIndex
        curPointIndex = 0
        
        'Next, we want to produce a similarly compatible list of opacity points.
        Dim numOpacityStops As Long
        
        Dim tmpOpacityList() As GradientPoint
        ReDim tmpOpacityList(0 To (m_GradientDefinition.gd_NumOpacityStops + 1) * 2) As GradientPoint
        
        For i = 0 To m_GradientDefinition.gd_NumOpacityStops - 1
            
            If (curPointIndex > UBound(tmpOpacityList)) Then ReDim Preserve tmpOpacityList(0 To curPointIndex * 2) As GradientPoint
            
            'First point
            If (i = 0) And (m_GradientDefinition.gd_Opacity(0).gc_Location <> 0) Then
                With m_GradientDefinition.gd_Opacity(i)
                    tmpOpacityList(curPointIndex).PointPosition = 0!
                    tmpOpacityList(curPointIndex).PointOpacity = .gc_Color.Alpha
                End With
                curPointIndex = curPointIndex + 1
            End If
            
            'Add the current point to the collection
            With m_GradientDefinition.gd_Opacity(i)
                tmpOpacityList(curPointIndex).PointOpacity = .gc_Color.Alpha
                tmpOpacityList(curPointIndex).PointPosition = .gc_Location / 100!
            End With
            
            curPointIndex = curPointIndex + 1
            
            'Midpoint
            If (i < m_GradientDefinition.gd_NumOpacityStops - 1) Then
                If (m_GradientDefinition.gd_Opacity(i).gc_Midpoint <> 50) Then
                    If (curPointIndex > UBound(tmpOpacityList)) Then ReDim Preserve tmpOpacityList(0 To curPointIndex * 2) As GradientPoint
                    
                    With m_GradientDefinition.gd_Opacity(i)
                        tmpA = .gc_Color.Alpha
                        tmpA = tmpA + Int(m_GradientDefinition.gd_Opacity(i + 1).gc_Color.Alpha)
                        tmpOpacityList(curPointIndex).PointOpacity = Int(tmpA / 2 + 0.5)
                        tmpPos = CSng(m_GradientDefinition.gd_Opacity(i + 1).gc_Location - .gc_Location) / 100!
                        tmpOpacityList(curPointIndex).PointPosition = (.gc_Location / 100!) + (tmpPos * (.gc_Midpoint / 100!))
                        curPointIndex = curPointIndex + 1
                    End With
                End If
            End If
            
            'Final point
            If (i = m_GradientDefinition.gd_NumOpacityStops - 1) And (m_GradientDefinition.gd_Opacity(i).gc_Location <> 100) Then
                If (curPointIndex > UBound(tmpOpacityList)) Then ReDim Preserve tmpOpacityList(0 To curPointIndex * 2) As GradientPoint
                With m_GradientDefinition.gd_Opacity(i)
                    tmpOpacityList(curPointIndex).PointOpacity = .gc_Color.Alpha
                    tmpOpacityList(curPointIndex).PointPosition = 1!
                End With
                curPointIndex = curPointIndex + 1
            End If
            
        Next i
        
        numOpacityStops = curPointIndex
        
        'We now have properly structured RGB and Opacity gradients.  The problem?  They're each on their
        ' own scales, and we need to mash 'em together.
        
        'There are myriad ways to optimize this, but as these collections aren't large, we do a
        ' simple variation on an insertion sort.
        
        'First, set the opacity of the first point in the color array
        tmpGradientList(0).PointOpacity = tmpOpacityList(0).PointOpacity
        
        'Next, set the opacity of the last point in the color array
        tmpGradientList(numColorStops - 1).PointOpacity = tmpOpacityList(numOpacityStops - 1).PointOpacity
        
        'Now we have to deal with all the values in-between those two.
        
        'First, perform a special check for uniform opacity.
        If (numOpacityStops = 2) Then
            If (tmpGradientList(0).PointOpacity = tmpGradientList(1).PointOpacity) Then
                For i = 0 To numColorStops - 1
                    tmpGradientList(i).PointOpacity = tmpGradientList(0).PointOpacity
                Next i
                GoTo OpacityHandled
            End If
        End If
        
        'If we're still here, opacity is not uniform across the gradient.  We need to manually merge
        ' opacity values with the underlying RGBA data (sigh).
        Dim j As Long, k As Long
        Dim opacityScaleFactor As Single, lastColorIndex As Long
        lastColorIndex = 0
        
        For i = 1 To numOpacityStops - 1
            
            'For each opacity stop, loop through the RGBA array until we find the first point
            ' with a position *after* the current one.
            For j = lastColorIndex To numColorStops - 1
                
                'If this point sits between the current opacity point and the previous one...
                If (tmpGradientList(j).PointPosition < tmpOpacityList(i).PointPosition) Then
                    
                    'Scale this point's opacity between the previous point and this one
                    opacityScaleFactor = (tmpOpacityList(i).PointPosition - tmpOpacityList(i - 1).PointPosition)
                    If (opacityScaleFactor <> 0!) Then opacityScaleFactor = (tmpGradientList(j).PointPosition - tmpOpacityList(i - 1).PointPosition) / opacityScaleFactor Else opacityScaleFactor = 1!
                    tmpGradientList(j).PointOpacity = tmpOpacityList(i).PointOpacity * opacityScaleFactor + (tmpOpacityList(i - 1).PointOpacity * (1! - opacityScaleFactor))
                
                'This point sits *ahead* of the current opacity point
                Else
                
                    'First, check equality.  It's easy because we can just assign opacity.
                    If (tmpGradientList(j).PointPosition = tmpOpacityList(i).PointPosition) Then
                        tmpGradientList(j).PointOpacity = tmpOpacityList(i).PointOpacity
                    
                    'Unequal.  We need to insert this opacity point as a new RGBA stop.
                    Else
                    
                        'Figure out what the RGBA value is at this opacity point's position
                        
                        '(Failsafe only)
                        If (j > 0) Then
                        
                            opacityScaleFactor = tmpGradientList(j).PointPosition - tmpGradientList(j - 1).PointPosition
                            If (opacityScaleFactor <> 0!) Then opacityScaleFactor = (tmpOpacityList(i).PointPosition - tmpGradientList(j - 1).PointPosition) / opacityScaleFactor Else opacityScaleFactor = 1!
                            If (opacityScaleFactor < 0!) Then opacityScaleFactor = 0!
                            If (opacityScaleFactor > 1!) Then opacityScaleFactor = 1!
                            
                            Dim newRGBA As RGBQuad
                            newRGBA.Blue = CSng(Colors.ExtractBlue(tmpGradientList(j).PointRGB)) * opacityScaleFactor + CSng(Colors.ExtractBlue(tmpGradientList(j - 1).PointRGB)) * (1! - opacityScaleFactor)
                            newRGBA.Green = CSng(Colors.ExtractGreen(tmpGradientList(j).PointRGB)) * opacityScaleFactor + CSng(Colors.ExtractGreen(tmpGradientList(j - 1).PointRGB)) * (1! - opacityScaleFactor)
                            newRGBA.Red = CSng(Colors.ExtractRed(tmpGradientList(j).PointRGB)) * opacityScaleFactor + CSng(Colors.ExtractRed(tmpGradientList(j - 1).PointRGB)) * (1! - opacityScaleFactor)
                            
                            'Create a new point representing what we need to insert
                            Dim tmpGradientPoint As GradientPoint
                            tmpGradientPoint.PointRGB = RGB(newRGBA.Red, newRGBA.Green, newRGBA.Blue)
                            tmpGradientPoint.PointPosition = tmpOpacityList(i).PointPosition
                            tmpGradientPoint.PointOpacity = tmpOpacityList(i).PointOpacity
                            
                            'Insert the new point into the collection
                            If (numColorStops > UBound(tmpGradientList)) Then ReDim Preserve tmpGradientList(0 To numColorStops * 2) As GradientPoint
                            For k = numColorStops To j + 1 Step -1
                                tmpGradientList(k) = tmpGradientList(k - 1)
                            Next k
                            tmpGradientList(j) = tmpGradientPoint
                            numColorStops = numColorStops + 1
                            lastColorIndex = j + 1
                            
                            'Resume with the next opacity point
                            Exit For
                            
                        'This is the very first point in the gradient list.  This case shouldn't trigger,
                        ' but if it does, simply copy the first gradient point into place.
                        Else
                            
                        End If
                            
                    End If
                    
                End If
                
            Next j
            
        Next i
        
        'RGB and Opacity data has been successfully merged!
        
OpacityHandled:
        
        'Start by loading all color data into a gradient object.
        Dim tmpGradient As pd2DGradient
        Set tmpGradient = New pd2DGradient
        tmpGradient.CreateGradientFromPointCollection numColorStops, tmpGradientList
        tmpGradient.SetGradientAngle m_GradientDefinition.gd_Angle
        tmpGradient.SetGradientShape P2_GS_Linear
        
        Select Case m_GradientDefinition.gd_Type
            Case keSGTLinear
                tmpGradient.SetGradientShape P2_GS_Linear
                If m_GradientDefinition.gd_Invert Then tmpGradient.ReverseGradient
            Case keSGTRectangular
                tmpGradient.SetGradientShape P2_GS_Rectangle
                If m_GradientDefinition.gd_Invert Then tmpGradient.ReverseGradient
            Case keSGTSunburst
                tmpGradient.SetGradientShape P2_GS_Radial
                If (Not m_GradientDefinition.gd_Invert) Then tmpGradient.ReverseGradient
            'Conical gradients are TODO; GDI+ doesn't support them natively
            Case keSGTRadial
        End Select
        
        'Previously "invert" was handled here, but note in the above select statement
        ' that its interpretation varies by shape!
        
        'For radial brushes, set center position
        If (m_GradientDefinition.gd_cX <> 50) Or (m_GradientDefinition.gd_cY <> 50) Then
            tmpGradient.SetCenterOffset m_GradientDefinition.gd_cX / 100!, m_GradientDefinition.gd_cY / 100!
        Else
            tmpGradient.SetCenterOffset_Default
        End If
        
        'Finally, create a matching gradient brush!
        If (dstBrush Is Nothing) Then Set dstBrush = New pd2DBrush
        dstBrush.SetBrushMode P2_BM_Gradient
        dstBrush.SetBrushGradientAllSettings tmpGradient.GetGradientAsString(True)
        
        okToProceed = psp_Success
    
    Else
        InternalError funcName, "null gradient brush chunk size", srcWarnings
        okToProceed = psp_Warning
        GoTo EarlyTermination
    End If
    
EarlyTermination:
    BuildBrush_Gradient = okToProceed
    
    Exit Function
    
'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError funcName, "internal VB error #" & Err.Number & ": " & Err.Description, srcWarnings
    srcWarnings.AddString "Internal error in pdPSPShape." & funcName & ", #" & Err.Number & ": " & Err.Description
    BuildBrush_Gradient = psp_Failure
    
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
    Dim i As Long, lastMoveTo As Long
    lastMoveTo = -1
    For i = 0 To m_NodeCount - 1
        
        'NOTE: I don't know what node visibility actually means, but skipping "invisible" nodes
        ' results in some shapes not being rendered, even though they're rendered when the file
        ' is loaded into Paintshop Pro.  As such, I've disabled this visibility check entirely.
        'If ((m_Nodes(i).pn_NodeFlags And keNodeVisible) = 0) Then GoTo ResumeNextPoint
        
        'When I first implemented this class, it was helpful to display node attributes.
        ' Here's the code in case you need to review parser behavior in the future:
        'PDDebug.LogAction i & ": " & m_Nodes(i).pn_MoveTo & " - " & m_Nodes(i).pn_X & ", " & m_Nodes(i).pn_Y & " - " & m_Nodes(i).pn_hX1 & ", " & m_Nodes(i).pn_hY1 & " - " & m_Nodes(i).pn_hX2 & ", " & m_Nodes(i).pn_hY2
        
        'On the first point, always start a new sub-path
        If (i = 0) Then
            m_PolylineShape.StartNewFigure
            If m_Nodes(i).pn_MoveTo Then lastMoveTo = i
            
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
                If (m_Nodes(i - 1).pn_MoveTo And (i <> 1)) Then
                    m_PolylineShape.StartNewFigure
                    lastMoveTo = i - 1
                End If
                
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
                If PSP_VECTOR_DEBUG_NO_CURVES Then
                    
                    m_PolylineShape.AddLine m_Nodes(i - 1).pn_X, m_Nodes(i - 1).pn_Y, m_Nodes(i).pn_X, m_Nodes(i).pn_Y
                    
                    'If this is a "closed" node, connect it to the last MoveTo node to complete this shape
                    If ((m_Nodes(i).pn_NodeFlags And keNodeClosed) <> 0) And (lastMoveTo >= 0) Then
                        m_PolylineShape.AddLine m_Nodes(i).pn_X, m_Nodes(i).pn_Y, m_Nodes(lastMoveTo).pn_X, m_Nodes(lastMoveTo).pn_Y
                    End If
                    
                Else
                    
                    m_PolylineShape.AddBezierCurve m_Nodes(i - 1).pn_X, m_Nodes(i - 1).pn_Y, m_Nodes(i - 1).pn_hX2, m_Nodes(i - 1).pn_hY2, m_Nodes(i).pn_hX1, m_Nodes(i).pn_hY1, m_Nodes(i).pn_X, m_Nodes(i).pn_Y
                    
                    'If this is a "closed" node, connect it to the last MoveTo node to complete this shape
                    If ((m_Nodes(i).pn_NodeFlags And keNodeClosed) <> 0) And (lastMoveTo >= 0) Then
                        m_PolylineShape.AddBezierCurve m_Nodes(i).pn_X, m_Nodes(i).pn_Y, m_Nodes(i).pn_hX2, m_Nodes(i).pn_hY2, m_Nodes(lastMoveTo).pn_hX1, m_Nodes(lastMoveTo).pn_hY1, m_Nodes(lastMoveTo).pn_X, m_Nodes(lastMoveTo).pn_Y
                    End If
                    
                End If
                
                'If this is the end of the current point collection, close the current path, since there
                ' are no more points to add.
                If (i = m_NodeCount - 1) Then m_PolylineShape.CloseCurrentFigure
                
            End If
            
        End If
        
ResumeNextPoint:
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

Private Sub Class_Initialize()
    
    'Debug purposes only
    Set m_Randomize = New pdRandomize
    m_Randomize.SetSeed_AutomaticAndRandom
    m_Randomize.SetRndIntegerBounds 0, 16777215
    
End Sub
