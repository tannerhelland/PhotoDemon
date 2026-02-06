VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdGlyphCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Glyph Collection Interface
'Copyright 2015-2026 by Tanner Helland
'Created: 27/April/15
'Last updated: 11/September/22
'Last update: add support for justified paragraph alignment
'
'This class manages the creation and storage of individual text glyphs.  These glyphs are used by PhotoDemon's
' advanced text tool to enable advanced text features.
'
'This class leans HEAVILY on pdUniscribe, PhotoDemon's comprehensive Uniscribe interface.  The relevant internal
' instance is m_Uniscribe.  Vector paths created by this class are ultimately rendered by a pdTextRenderer instance.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Copy of the GDI font handle used for glyph generation.
' IMPORTANT NOTE: this class will select the font handle into a DC of its own creation.
' Prior to passing a font handle to this class, *YOU MUST UNSELECT IT FROM ANY DCs*.
Private m_GDIFont As Long

'TrueType fonts use their own special fraction format.  Yes, it's obnoxious.
Private Type TT_FIXED
    Fract As Integer
    IntValue As Integer
End Type

'The MAT2 type consists of 4 FIXED values (16-bits of fraction (unsigned), followed by 16-bits of integer (signed))
' that comprise the four top-left spots of a 3x3 matrix.  We do actually use this matrix to reorient characters
' against their top-left corner.
Private Type MAT2
    eM11 As TT_FIXED
    eM12 As TT_FIXED
    eM21 As TT_FIXED
    eM22 As TT_FIXED
End Type

'Supported shape types inside TrueType and OpenType fonts
Private Const TT_PRIM_LINE As Long = 1       'Line or polyline primitive
'Private Const TT_PRIM_QSPLINE As Long = 2    'Quadratic Bezier spline, not used by PD at present
                                              ' (because we forcibly request cubic splines for simplicity)
Private Const TT_PRIM_CSPLINE As Long = 3    'Cubic Bezier spline

'This GLYPHMETRICS struct is passed to GetGlyphOutline; it contains basic positioning data, separate from the glyph itself.
' Note that all measurements are in *device units*.
Private Type GLYPHMETRICS
    gmBlackBoxX As Long    'Width and height of the smallest rectangle that completely encloses the glyph (its "black box").
    gmBlackBoxY As Long
    gmptGlyphOrigin As PointLong  '(X, Y) coordinates of the upper left corner of the smallest rectangle
                                  ' that completely encloses the glyph.
    gmCellIncX As Integer   'Horizontal and vertical distance from the origin of the current character cell
                            ' to the origin of the next character cell.
    gmCellIncY As Integer   'IMPORTANT NOTE: these values do not include kerning.  (How could they,
                            ' when they're generated for a standalone glyph!)
End Type

'Generic GDI object management
Private m_oldFontHandle As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Glyph retrieval.  Per MSDN (https://msdn.microsoft.com/en-us/library/dd144891%28v=vs.85%29.aspx),
' I don't think there's any difference between the A and W variants, but W is declared for consistency
' with the rest of PD.
Private Const GGO_BEZIER As Long = 3
Private Const GGO_GLYPH_INDEX As Long = &H80&
Private Const GGO_UNHINTED As Long = &H100&
Private Declare Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineW" (ByVal hDC As Long, ByVal uChar As Long, ByVal uFormat As Long, ByRef lpgm As GLYPHMETRICS, ByVal cbBuffer As Long, ByVal ptrToBuffer As Long, ByRef lpmat2 As MAT2) As Long

'This type is the minimum amount of data required to properly store and render a given glyph outline.
' The pd2DPath instance will be filled with a corresponding GDI+ GraphicsPath, assuming translation was successful.
Private Type pdGlyphOutline
    glyphIndex As Long        'Index of the current glyph inside this font.  This is not guaranteed to have a
                              ' 1:1 mapping to character codes!  Note also that PD no longer stores character
                              ' code values, as they are not relevant during rendering.
    glyphPath As pd2DPath     'GDI+ GraphicsPath wrapper containing the fully translated font outline
    gMetrics As GLYPHMETRICS  'Copy of the GLYPHMETRICS struct returned by GetGlyphMetrics
    hasOutline As Boolean     'If an outline is successfully retrieved and parsed, this will be set to TRUE.
                              ' If no rendering is required (whitespace, control chars, etc), this will be FALSE.
End Type

'Current glyph outline collection.  Because glyph generation is very expensive, we store glyphs as we translate them.
' If the font hasn't changed, we can reuse glyphs from past calculations.
Private m_GlyphOutlines() As pdGlyphOutline

'New glyph collection format, using Uniscribe for generation
Private m_Glyphs() As PDGlyphUniscribe

'Current number of generated glyphs.  This does not correspond to the UBound() of the glyph collection.
Private m_NumOfGlyphs As Long
Private Const INITIAL_GLYPH_COLLECTION_SIZE As Long = 128

'Current number of generated glyph outlines.  This does not correspond to the UBound() of the glyph outline collection,
' or to m_NumOfGlyphs, above.  PD only stores one copy of each UNIQUE glyph outline, whereas m_Glyphs() stores positioning
' data on each occurrence of each glyph.
Private m_NumOfGlyphOutlines As Long

'Generic identity matrix, as required by GetGlyphOutline
Private m_IdentityMatrix As MAT2

'Tiny temporary DIB, created once and cached; the target font is selected into this DIB prior to retrieving glyph outlines.
Private m_tmpDIB As pdDIB

'Uniscribe interface.  Uniscribe is used for initial script itemizing, shaping, layout and placement, but note that
' we further modify the Uniscribe output to apply custom PhotoDemon features (like line-spacing, jitter, etc)
Private m_Uniscribe As pdUniscribe

'Metrics for the current font.  These are crucial because font path coordinate are all relative to the font baseline,
' but when rendering a path to the screen, we need to orient the points against the top-left corner of the image.
' It's faster to apply this coordinate conversion when we first construct the glyphs.
Private m_TextMetrics As TEXTMETRIC
Private m_OutlineTextMetrics As OUTLINETEXTMETRIC

'During glyph parsing, we have to apply a number of fixed conversions due to differences between TrueType coordinates
' and traditional screen coordinates.  We cache these values so we don't have to constantly retrieve them from the
' larger TEXTMETRIC struct.
Private m_FontAscent As Long, m_FontDescent As Long, m_FontHeight As Long

'Copy of the string, font, and DC used by previous runs.  If any of these change, we have to repeat various
' Uniscribe steps, but we can shortcut some steps if (for example) font setting changes but the target string does not.
Private m_LastFont As Long, m_LastDC As Long, m_LastString As String

'When setting alignment, we improve performance in the inner loop by calculating running line widths as we go.
' To prevent churn, we use a module-level array and only reset it as necessary.
Private m_LineWidths() As Single
Private Const DEFAULT_INITIAL_LINE_COUNT As Long = 16

'This class supports various custom layout and character settings, which the caller can supply prior to final path assembly.
Private m_LineSpacing As Single
Private m_CharSpacing As Single
Private m_CharOrientation As Single
Private m_CharJitterX As Single
Private m_CharJitterY As Single
Private m_CharInflation As Single
Private m_CharMirror As PD_CharacterMirror

'If no custom character modifications are active, we can skip ceratin steps during path assembly.
Private m_CharEffectsEnabled As Boolean

'Underline and strikeout data is not embedded in fonts; these features must be manually rendered
Private m_StyleUnderline As Boolean, m_StyleStrikeout As Boolean

'Prior to glyph generation, this class needs a copy of GDI font handle.  I have deliberately made this
' function separate AS A REMINDER TO REMOVE THE FONT FROM ANY DCs PRIOR TO GIVING THIS CLASS ACCESS.
Friend Sub NotifyOfGDIFontChoice(ByVal srcGDIFontHandle As Long, Optional ByVal underlineActive As Boolean = False, Optional ByVal strikeoutActive As Boolean = False)

    'If this font matches previous font requests, we can retain existing glyph information.
    
    'If it doesn't match, we need to start over from scratch.  Alas.
    If (m_GDIFont <> srcGDIFontHandle) Then
        
        'Wipe the current glyph collection
        ResetGlyphCollection
        
        'Copy the font handle, then notify Uniscribe that it needs to generate a new script cache
        m_GDIFont = srcGDIFontHandle
        m_Uniscribe.UpdateFont
        
    End If
    
    'Select the font into our temporary DC.  (Many GDI functions operate on DCs instead of hFonts, so this is important.)
    m_oldFontHandle = SelectObject(m_tmpDIB.GetDIBDC, m_GDIFont)
    
    'Retrieve basic text metrics for this font, and cache any metrics required on the inner glyph parsing loop
    Fonts.FillTextMetrics m_tmpDIB.GetDIBDC, m_TextMetrics
    m_FontAscent = m_TextMetrics.tmAscent
    m_FontDescent = m_TextMetrics.tmDescent
    m_FontHeight = m_TextMetrics.tmHeight
    
    'Also retrieve outline text metrics, which include basic text metrics plus a whole bunch more.
    Fonts.FillOutlineTextMetrics m_tmpDIB.GetDIBDC, m_OutlineTextMetrics
    
    'Cache the underline and strikeout values as well
    m_StyleUnderline = underlineActive
    m_StyleStrikeout = strikeoutActive
        
End Sub

'After this class has done its work, the caller MUST CALL THIS FUNCTION to release their font from our internal DC
Friend Sub RequestGDIFontRelease()
    SelectObject m_tmpDIB.GetDIBDC, m_oldFontHandle
End Sub

'Given a source string, add all glyphs from that string to our current collection
Friend Function BuildGlyphCollection(ByRef srcString As String, Optional ByVal useHinting As Boolean = False, Optional ByVal remapType As PD_StringRemap = sr_None) As Boolean
    
    Dim curGlyphIndex As Long
    Dim charAlreadyExists As Boolean
    Dim rawGlyphBytes() As Long
    Dim glyphAdded As Boolean
    
    'If an optional string remap is specified, apply it now.
    If (remapType <> sr_None) Then srcString = Strings.StringRemap(srcString, remapType)
    
    Dim stringChanged As Boolean, fontOrDCChanged As Boolean
    stringChanged = Strings.StringsNotEqual(srcString, m_LastString, vbBinaryCompare)
    fontOrDCChanged = (m_LastFont <> m_GDIFont) Or (m_tmpDIB.GetDIBDC <> m_LastDC)
    
    'If the string has changed since the last call, we must re-itemize it.
    If stringChanged Then
    
        'Use Uniscribe to process a list of required glyphs and their relevant positions.
        m_Uniscribe.Step1_ScriptItemize srcString
        m_Uniscribe.Step2_ScriptLayout
        m_LastString = srcString
        
    End If
    
    'If the string *or* font has changed, we must recreate our glyph collection
    If stringChanged Or fontOrDCChanged Then
    
        'Make a note of the string contents; if they don't change on the next glyph collection call, we can skip these steps
        m_Uniscribe.Step3_ScriptShape m_tmpDIB.GetDIBDC
        m_Uniscribe.Step4_ScriptPlace m_tmpDIB.GetDIBDC
        m_Uniscribe.Step5_ScriptBreak
        
        'Make a note of the font and DC used
        m_LastFont = m_GDIFont
        m_LastDC = m_tmpDIB.GetDIBDC
        
        'Retrieve a copy of Uniscribe's data in a PD-friendly format.
        m_NumOfGlyphs = m_Uniscribe.GetCopyOfGlyphCache(m_Glyphs, m_tmpDIB.GetDIBDC)
        
        'Curious about what pdUniscribe returns?  Uncomment this line for a nice intro.
        'm_Uniscribe.printUniscribeDebugInfo
        
        'Start iterating through the glyph collection returned by Uniscribe.  For any glyph indices that DO NOT EXIST
        ' in our current outline collection, retrieve a raw outline buffer and convert it to a pd2DPath object.
        '
        'IMPORTANT NOTE!  The glyph collection and glyph outline collections do *NOT* have a 1:1 mapping.
        ' Because glyph retrieval is expensive, we only store one copy of each unique glyph outline.
        ' So for the text AAAaaa, the m_Glyphs collection will have six entries (one for each glyph),
        ' while the m_GlyphOutline collection will have two entries (one for each *UNIQUE* glyph).
        ' Mapping between the two arrays is done by matching .GlyphIndex values in the respective structs.
        Dim i As Long, j As Long
        For i = 0 To m_NumOfGlyphs - 1
            
            charAlreadyExists = False
            glyphAdded = False
            curGlyphIndex = m_Glyphs(i).glyphIndex
            
            'See if this glyph index already exists in our outline collection
            If (m_NumOfGlyphs > 0) Then
                
                For j = 0 To m_NumOfGlyphOutlines - 1
                    
                    'If this glyph already exists, mark it and exit immediately
                    If curGlyphIndex = m_GlyphOutlines(j).glyphIndex Then
                        charAlreadyExists = True
                        Exit For
                    End If
                    
                Next j
                
            End If
            
            'If the current character isn't in our collection, retrieve it now
            If (Not charAlreadyExists) Then
                
                'Mark the glyph index (index into the *font file*), so we can match individual glyphs
                ' in the string to this outline path.
                m_GlyphOutlines(m_NumOfGlyphOutlines).glyphIndex = curGlyphIndex
                
                'Retrieve the raw glyph buffer.  This step will fail on non-display characters (e.g. control chars,
                ' like linefeeds, as well as whitespace characters like spaces or tabs), so failure is *not a
                ' bad thing*.  What makes handling tricky is that failure can occur at many different places;
                ' the function itself may fail, or it may succeed but return a zero-length buffer, or it may
                ' succeed but return a blank outline (an empty square or square with an x through it).  We don't care
                ' *why* it fails - we only care if this stage produces a useable outline or not.
                If GetRawGlyphBuffer(curGlyphIndex, rawGlyphBytes, useHinting) Then
                    
                    'With the raw buffer retrieved, we can now walk the byte array and retrieve individual line
                    ' and curve metrics.
                    
                    'Before doing that, see if the buffer size is non-zero.  A zero-size buffer is valid output
                    ' for whitespace characters (e.g. spaces).
                    If (UBound(rawGlyphBytes) > 0) Then
                        
                        'Debug.Print ChrW$(curChar) & ":" & UBound(rawGlyphBytes)
                        
                        'Proceed with parsing this glyph
                        glyphAdded = ParseRawGlyphBuffer(rawGlyphBytes, m_GlyphOutlines(m_NumOfGlyphOutlines).glyphPath)
                        m_GlyphOutlines(m_NumOfGlyphOutlines).hasOutline = glyphAdded
                        
                        'Debug only:
                        'If (Not glyphAdded) Then PDDebug.LogAction "WARNING!  parseRawGlyphBuffer failed for glyph index " & curGlyphIndex
                        
                    'This is likely a control character.  We still add it (to avoid attempted parsing on future attempts),
                    ' but we specifically note that it doesn't need to be rendered to the screen.
                    Else
                        'Debug.Print "Zero-length glyph buffer: " & curGlyphIndex & ":" & UBound(rawGlyphBytes)
                        glyphAdded = True
                        m_GlyphOutlines(m_NumOfGlyphOutlines).hasOutline = False
                    End If
                
                'Raw glyph buffer retrieval failed.  This is a valid return for some control characters.
                ' To avoid future attempts at parsing an outline, we *still* add it to the collection,
                ' but we note that it isn't renderable.
                Else
                    'Debug.Print "getRawGlyphBuffer full-on failed: " & curGlyphIndex
                    m_GlyphOutlines(m_NumOfGlyphOutlines).hasOutline = False
                    glyphAdded = True
                End If
                
                'If a glyph was successfully added to the collection, increment the current glyph count and resize
                ' the collection as necessary.
                If glyphAdded Then
                    m_NumOfGlyphOutlines = m_NumOfGlyphOutlines + 1
                    If (m_NumOfGlyphOutlines > UBound(m_GlyphOutlines)) Then ReDim Preserve m_GlyphOutlines(0 To m_NumOfGlyphOutlines * 2 - 1) As pdGlyphOutline
                End If
                
            '(End "If Not charAlreadyExists")
            End If
            
        Next i
        
    End If
    
End Function

'Given a Unicode entry point, fill a buffer with the raw output of GetGlyphOutline
Private Function GetRawGlyphBuffer(ByVal curGlyphIndex As Long, ByRef dstBuffer() As Long, Optional ByVal useHinting As Boolean = False) As Boolean
    
    'We could start by retrieving glyph metrics, which are plugged directly into the current
    ' glyph entry in the central collection.  (These metrics are separate from the glyph shape;
    ' we use them to control inter-character spacing, among other things.)
    
    'To grab these, use the following line of code:
    'ggoReturn = GetGlyphOutline(m_tmpDIB.getDIBDC, curChar, GGO_METRICS, m_glyphoutlines(m_NumOfGlyphs).gMetrics, 0, 0, m_IdentityMatrix)
    
    'The reason I don't do this is that the metrics will be retrieved anyway by our next call,
    ' which requests length of the raw glyph buffer.  For performance reasons, we kill two birds
    ' with one stone on that call.
    
    'Prior to retrieving metrics, set the hinting flag.  Note that hinting is optional, and its
    ' presence (or lack thereof) has consequences for glyph shape and positioning.
    Dim formatFlags As Long
    formatFlags = GGO_BEZIER Or GGO_GLYPH_INDEX
    If (Not useHinting) Then formatFlags = formatFlags Or GGO_UNHINTED
    
    'Next, we want to retrieve the required size of the glyph shape buffer.  FYI, for complex glyphs,
    ' this can be fairly large.  Note also that we explicitly request bezier curve format.
    ' This makes our subsequent parsing much easier, and improves the output fidelity of OpenType fonts.
    Dim ggoReturn As Long, byteSizeOfBuffer As Long
    ggoReturn = GetGlyphOutline(m_tmpDIB.GetDIBDC, curGlyphIndex, formatFlags, m_GlyphOutlines(m_NumOfGlyphOutlines).gMetrics, 0, 0, m_IdentityMatrix)
    
    'ggoReturn should be > 0, which tells us the required size of the destination buffer.
    If (ggoReturn > 0) Then
        byteSizeOfBuffer = ggoReturn
        
    'If zero is returned, that's okay; this might just be a whitespace character with no corresponding glyph.
    ' Return TRUE; the caller function has logic to deal with this combination of outputs.
    Else
        ReDim dstBuffer(0) As Long
        GetRawGlyphBuffer = True
        Exit Function
    End If
    
    'Prep the buffer.  For convenience, I've declared the buffer as type LONG to prevent doing a lot of
    ' manual parsing.  This size strategy requires us to divide the byte size of the buffer by 4 (obviously).
    ReDim dstBuffer(0 To (byteSizeOfBuffer \ 4) - 1) As Long
    
    'Now we can retrieve the actual data buffer!
    ggoReturn = GetGlyphOutline(m_tmpDIB.GetDIBDC, curGlyphIndex, formatFlags, m_GlyphOutlines(m_NumOfGlyphOutlines).gMetrics, byteSizeOfBuffer, VarPtr(dstBuffer(0)), m_IdentityMatrix)
    
    'Per MSDN, successful queries return a value > 0
    GetRawGlyphBuffer = (ggoReturn > 0)
    
End Function

'Given the raw byte stream returned by GetGlyphOutline, walk the buffer and assemble a matching GDI+ GraphicsPath
' from the output.
Private Function ParseRawGlyphBuffer(ByRef srcBuffer() As Long, ByRef dstGlyphPath As pd2DPath) As Boolean
    
    'Parsing the output of GetGlyphOutline isn't quite as nasty as you'd thing.  The structs themselves are
    ' pretty straightforward; it's the asinine data types used that make things unpleasant.
    
    'This MSDN article provides an incredibly helpful overview.  (That might be the first time I've ever said this.)
    ' https://support.microsoft.com/en-us/kb/243285
    
    'First, you need to understand a few crucial structs.  We don't actually declare these structs
    ' (because copying data to/from them is a giant waste of time when we can just read the bytes directly).
    ' These structs are:
    
    'Type POINTFX
    '    FIXED x
    '    FIXED y
    'End Type
    
    'FIXED is a 4-byte type with 2 bytes of unsigned fraction ( / 65536), followed by 2 bytes of signed integer.
    ' From the relevant Dr Dobbs article (http://www.drdobbs.com/truetype-font-secrets/184403680):
    '
    ' "Value represents the part of the real number to the left of the decimal point; fract represents the part
    '  to the right of the decimal point, considered as a fraction of 65,536. For example, 0.5 becomes
    '  (fract, value) = (32768, 0); 2.25 is equivalent to (fract, value) = (16384, 2)."
    
    'Type TTPOLYGONHEADER
    '    cb As Long            'Bytes required by the TTPOLYGONHEADER structure and any TTPOLYCURVE structure(s) that follow it.
    '    dwType As Long        'Always the TT_POLYGON_TYPE constant, so we can happily ignore it
    '    pfxStart As POINTFX   'Starting point of the contour in the character outline.
    'End Type
    
    'Each TTPOLYGONHEADER is followed by one or more TTPOLYCURVE structs.  The number of structs can't be
    ' inferred in advance, because curve size varies according to complexity.
    
    'Type TTPOLYCURVE
    '    wType As Integer    'Type of the curve.  This can be one of three values; see below
    '    cpfx As Integer     'The number of POINTFX structures in this curve (e.g. the size of the curve array,
    '                         if you want to approach it that way)
    '    apfx() As POINTFX   'Array of POINTFX structures, with number of entries = cpfx
    'End Type
    
    'TTPOLYCURVES take three forms, as marked by wType:
    ' 1) Line and/or polyline primitives (TT_PRIM_LINE)
    ' 2) Quadratic Bezier splines (TT_PRIM_QSPLINE).  This is the default curve format for TrueType fonts.
    ' 3) Cubic Bezier splines (TT_PRIM_CSPLINE).  This is the default curve format for OpenType fonts, but to get this format,
    '                                              we have to explicitly specify the GGO_BEZIER flag.  (PD does exactly this,
    '                                              to preserve better OpenType glyph shaping.)
    
    'These structs are all fairly straightforward; the biggest hit is converting the damn POINTFX structs
    ' into usable floats.

    'The raw byte stream handed to us by GetGlyphOutline can be read as follows:
    ' 1) A glyph outline is returned as a series of one or more contours defined by a TTPOLYGONHEADER structure
    '    followed by one or more curves.
    ' 2) Each curve in the contour is defined by a TTPOLYCURVE structure followed by a number of POINTFX data points.
    ' 3) POINTFX points are absolute positions, not relative moves.  (This is awesome, actually, as it makes GDI+ conversion
    '    *much* easier)
    ' 4) The starting point of a contour is given by the pfxStart member of the TTPOLYGONHEADER structure.
    ' 5) The starting point of each curve is the last point of the previous curve or the starting point of the contour.
    ' 6) The count of data points in a curve is stored in the cpfx member of TTPOLYCURVE structure.
    ' 7) The size of each contour in the buffer, in bytes, is stored in the cb member of TTPOLYGONHEADER structure.
    ' 8) Additional curve definitions are packed into the buffer following preceding curves, and additional contours
    '    are packed into the buffer following preceding contours.
    ' 9) The buffer contains as many contours as fit within the buffer returned by GetGlyphOutline.
    
    'This seems like a lot of caveats, but honestly, it's not too bad to walk the struct like this.  Let's begin!
    
    'Start by initializing the destination graphics path
    Set dstGlyphPath = New pd2DPath
    
    Dim firstPoint As PointFloat, prevPoint As PointFloat, nextPoint As PointFloat
    Dim bezPoint2 As PointFloat, bezPoint3 As PointFloat
    Dim curveType As Long, pointCount As Long
    Dim i As Long
    
    Dim curPosition As Long
    curPosition = 0
    
    Dim endOfCurrentCurve As Long
    endOfCurrentCurve = 0
    
    Do
        
        'Start by parsing this TTPOLYGONHEADER, which defines a single contiguous shape.
        
        'Type TTPOLYGONHEADER
        '    cb As Long            'Bytes required by this TTPOLYGONHEADER structure and any structure(s) that follow it.
        '    dwType As Long        'Always the TT_POLYGON_TYPE constant, so we can happily ignore it
        '    pfxStart As POINTFX   'Starting point of this contour in the character outline.
        'End Type
        
        'Technically, we could start this task by double-checking the required TTPOLYGONHEADER.dwType value,
        ' which should always be TT_POLYGON_TYPE... but since it never deviates, let's just ignore it.
        ' (FYI it would occur at position "curPosition + 1")
        
        'First things first: retrieve the ending position of this struct.  We'll use this to know when to close
        ' the current curve. (Remember that buffer sizes are in bytes, but our array is Long-type.)
        endOfCurrentCurve = curPosition + srcBuffer(curPosition) \ 4
        
        'Next, retrieve the first point in the glyph.  Note that this point is relative to the origin of the glyph,
        ' and the origin of a glyph is *always* the lower left corner of the character at the character's baseline.
        firstPoint.x = GetSingleFromFIXED(srcBuffer(curPosition + 2))
        firstPoint.y = GetSingleFromFIXED(srcBuffer(curPosition + 3))
        
        'Also note the first point as the "previous point"; this simplifies our calculations on the inner loop
        prevPoint.x = firstPoint.x
        prevPoint.y = firstPoint.y
        
        'Advance the buffer pointer
        curPosition = curPosition + 4
        
        'With the first point determined, we can now proceed with parsing the next shape in line.
        Do
        
            'Next, we parse a TT_POLYCURVE struct, which looks like this:
            
            'Type TTPOLYCURVE
            '    wType As Integer    'Type of the curve.  This can be one of three values; see below
            '    cpfx As Integer     'The number of POINTFX structures in this curve (e.g. the UBound of the curve array,
            '                         if you want to approach it that way)
            '    apfx() As POINTFX   'Array of POINTFX structures, with number of entries = cpfx
            'End Type
            
            'Start by retrieving the curve type and number of points in this curve
            curveType = srcBuffer(curPosition) And 65535
            pointCount = srcBuffer(curPosition) \ 65536
            
            'Advance the pointer to the start of the POINTFX array
            curPosition = curPosition + 1
            
            Select Case curveType
            
                'Polyline
                Case TT_PRIM_LINE
                
                    'Polylines are easiest: just iterate each line, adding segments as we go
                    For i = 1 To pointCount
                        
                        'Add this line segment
                        nextPoint.x = GetSingleFromFIXED(srcBuffer(curPosition))
                        nextPoint.y = GetSingleFromFIXED(srcBuffer(curPosition + 1))
                        dstGlyphPath.AddLine prevPoint.x, prevPoint.y, nextPoint.x, nextPoint.y
                        
                        'Copy the next point into the previous point marker
                        prevPoint = nextPoint
                        
                        'Advance to the next buffer position
                        curPosition = curPosition + 2
                        
                    Next i
                
                'Quadratic Bezier spline (PD doesn't currently handle this case, because we explicitly request
                ' cubic bezier splines from the font mapper).
                'Case TT_PRIM_QSPLINE
                
                'Cubic Bezier spline
                Case TT_PRIM_CSPLINE
                    
                    'Cubic splines occur in sets of four points (but one is already known from the end of
                    ' the previous curve).  Iterate each in turn.
                    For i = 1 To pointCount Step 3
                    
                        'Previous point is already known; retrieve the next three points explicitly
                        bezPoint2.x = GetSingleFromFIXED(srcBuffer(curPosition))
                        bezPoint2.y = GetSingleFromFIXED(srcBuffer(curPosition + 1))
                        
                        bezPoint3.x = GetSingleFromFIXED(srcBuffer(curPosition + 2))
                        bezPoint3.y = GetSingleFromFIXED(srcBuffer(curPosition + 3))
                        
                        nextPoint.x = GetSingleFromFIXED(srcBuffer(curPosition + 4))
                        nextPoint.y = GetSingleFromFIXED(srcBuffer(curPosition + 5))
                        
                        dstGlyphPath.AddBezierCurve prevPoint.x, prevPoint.y, bezPoint2.x, bezPoint2.y, bezPoint3.x, bezPoint3.y, nextPoint.x, nextPoint.y
                        
                        'Copy the next point into the previous point marker
                        prevPoint = nextPoint
                        
                        'Advance to the next buffer position
                        curPosition = curPosition + 6
                    
                    Next i
            
            End Select
        
        'If there are more lines and/or curves in this shape, continue on with the next one
        Loop Until curPosition >= endOfCurrentCurve
    
        'This shape is now complete.  Close the path.
        dstGlyphPath.CloseCurrentFigure
    
    'If more shapes exist, draw them next
    Loop Until curPosition >= UBound(srcBuffer)
    
    'With parsing complete, we now need to translate the finished glyph downward.  TrueType and OpenType fonts
    ' are positioned relative to their baseline, so (0, 0) is the bottom-left point of the glyph that touches
    ' the baseline.  When rendering to the screen, we obviously want things oriented against their top-left corner.
    
    'We can do this by translating the path downward by the current font's ascender value, which is constant for all glyphs.
    dstGlyphPath.TranslatePath 0, m_FontAscent
    
    'Parsing complete!
    ParseRawGlyphBuffer = True

End Function

'Prior to calling AssembleCompositePath, below, the caller can notify us of custom layout settings via this function
Friend Sub NotifyCustomLayoutSettings(Optional ByVal customLineSpacing As Single = 0!, Optional ByVal customCharSpacing As Single = 0!, Optional ByVal customCharOrientation As Single = 0!, Optional ByVal customCharJitterX As Single = 0!, Optional ByVal customCharJitterY As Single = 0!, Optional ByVal customCharInflation As Single = 0!, Optional ByVal customCharMirror As PD_CharacterMirror = cm_None)
    
    m_LineSpacing = customLineSpacing
    m_CharSpacing = customCharSpacing
    m_CharOrientation = customCharOrientation
    m_CharJitterX = customCharJitterX
    m_CharJitterY = customCharJitterY
    m_CharInflation = customCharInflation
    m_CharMirror = customCharMirror
    
    'If one or more custom character attributes are non-zero, mark a class-level flag.
    ' This allows us to skip some processor-intensive steps when text uses normal formatting.
    m_CharEffectsEnabled = (m_CharSpacing <> 0) Or (m_CharOrientation <> 0) Or (m_CharJitterX <> 0) Or (m_CharJitterY <> 0) Or (m_CharInflation <> 0) Or (m_CharMirror <> cm_None)
    
End Sub

'After assembling a full glyph collection, this function can be called to generate a totally complete graphics path,
' with all characters laid out according to the passed rect.
Friend Function AssembleCompositePath(ByRef dstPath As pd2DPath, ByRef boundingRect As RectF, ByVal horizontalAlignment As GP_StringAlignment, ByVal verticalAlignment As GP_StringAlignment, Optional ByVal lineWrapMode As PD_TextWordwrap = tww_AutoWord, Optional ByVal stretchToFit As PD_TextStretchToFit = stf_None, Optional ByVal justifyLastLine As GP_StringAlignment = StringAlignmentNear) As Boolean
    
    'Initialize the destination path as necessary
    If (dstPath Is Nothing) Then Set dstPath = New pd2DPath Else dstPath.ResetPath
    
    'Only proceed if the source string is non-empty
    If (m_NumOfGlyphs > 0) Then
        
        'Path assembly happens in two steps.
        ' 1) First, precalculate all glyph (x, y) positions, including line-breaking.  Calculating positions in advance
        '    *greatly* simplifies the main rendering loop.  Note that some glyph-specific features (e.g. stretching)
        '    must be calculated here, to ensure line-breaks are placed correctly.
        ' 2) Once individual positions are calculated on a per-glyph basis, combine all glyphs into a single merged
        '    pd2DPath object.  Any glyph-specific features (e.g. mirroring, per-char rotation, etc) can be applied
        '    immediately prior to adding a glyph to the composite path.
        
        '(FYI: the first step - the glyph positioning algorithm - fills the .finalX and .finalY entries for each glyph.)
        
        'It's a bit counterintuitive, but the precalculation step also fills underline and strikeout paths for us.
        ' Historically, we computed these right here in this function, but that's far more complicated as it requires
        ' line-level data that CalculateGlyphPositions() does not return (e.g. identifying whitespace characters that
        ' sit at the end of a line, which should be ignored when displaying underline/strikeout).  As such, we simply
        ' pass blank path objects to CalculateGlyphPositions(), and merge the results with our final path when underline
        ' or strikeout styles are active.
        Dim additionalStylesRect As pd2DPath
        CalculateGlyphPositions boundingRect, horizontalAlignment, verticalAlignment, lineWrapMode, additionalStylesRect, stretchToFit, justifyLastLine
        
        Dim i As Long, curGlyph As Long, curGlyphOutline As Long
        Dim tmpGlyphOutlineCopy As pd2DPath, tmpMatrix As pd2DTransform, cRandom As pdRandomize
        Dim maxJitterX As Single, maxJitterY As Single, halfJitterX As Single, halfJitterY As Single
        Dim glyphIsRenderable As Boolean
                
        'If character-level orientation changes are active, prep some reusable transformation objects
        If m_CharEffectsEnabled Then
        
            Set tmpGlyphOutlineCopy = New pd2DPath
            Set tmpMatrix = New pd2DTransform
            
            'Seeding the random number generator is a little tricky.  We need it to be deterministic,
            ' so the layer's appearance doesn't change on save/load events, but we also need it to vary
            ' by changes to jitter values.
            Set cRandom = New pdRandomize
            cRandom.SetSeed_Int m_CharJitterX * 100 + m_CharJitterY * 10000
            
            maxJitterX = m_TextMetrics.tmAveCharWidth * (m_CharJitterX / 100)
            halfJitterX = maxJitterX / 2
            
            maxJitterY = m_TextMetrics.tmHeight * (m_CharJitterY / 100)
            halfJitterY = maxJitterY / 2
            
        End If
        
        'Start iterating through the source string one character at a time.
        For i = 0 To m_NumOfGlyphs - 1
            
            'Start by retrieving this glyph index
            curGlyph = i
            curGlyphOutline = GetGlyphOutlineIndex(curGlyph)
            
            'A glyph is renderable only if it...
            ' 1) was successfully processed by GetGlyphOutline, and...
            ' 2) GetGlyphOutline returned a shape with at least one line or curve
            glyphIsRenderable = (curGlyphOutline >= 0) And m_GlyphOutlines(curGlyphOutline).hasOutline
            
            'We also ignore things like spaces and control characters
            If glyphIsRenderable Then
                If (m_Glyphs(i).isHardLineBreak) Or (m_Glyphs(i).isWhiteSpace) Or (m_Glyphs(i).isZeroWidth) Then
                    glyphIsRenderable = False
                End If
            End If
            
            'If these characteristics are met, add this path to the merged path, at the offsets pre-calculated
            ' by CalculateGlyphPositions().
            If glyphIsRenderable Then
                
                'If weird character effects are active, we need to do a bit more processing work
                If m_CharEffectsEnabled Then
                
                    'Start by making a copy of the original glyph outline
                    tmpGlyphOutlineCopy.CloneExistingPath m_GlyphOutlines(curGlyphOutline).glyphPath
                    
                    'Let's break down the possible character changes one at a time.
                    
                    'm_CharOrientation controls character orientation, which is used to rotate characters around
                    ' their own center point.
                    If (m_CharOrientation <> 0) Then tmpGlyphOutlineCopy.RotatePathAroundItsCenter m_CharOrientation
                    
                    'm_CharMirror reflects characters across their own central axis
                    If (m_CharMirror <> 0) Then
                        
                        Select Case m_CharMirror
                        
                            Case cm_Horizontal
                                tmpGlyphOutlineCopy.MirrorPathAroundItsCenter True, False
                            
                            Case cm_Vertical
                                tmpGlyphOutlineCopy.MirrorPathAroundItsCenter False, True
                            
                            Case cm_Both
                                tmpGlyphOutlineCopy.MirrorPathAroundItsCenter True, True
                        
                        End Select
                        
                    End If
                    
                    'm_CharInflation widens the character paths using the generic GDI+ path widening algorithm
                    If (m_CharInflation <> 0) Then
                        
                        'So, new series of steps.
                        
                        'First, we want to flatten this path (e.g. convert it to a series of lines *only*,
                        ' no curves).  If we don't do this manually, GDI+ will do it for us.
                        
                        'TODO
                        
                        'Next, divide the path into standalone subpaths (necessary for letters like "o",
                        ' which have holes cut out).  Each subpath will be individually optimized.
                        
                        'TODO
                        
                        'Next, optimize each subpath by removing segments too small to see on a display
                        
                        'TODO
                        
                        'Finally, reassemble the original glyph by combining all subpaths into a single
                        ' "merged" path object.
                        
                        'TODO
                        
                        'Calculate the outline of the optimized path, as drawn by a pen with width m_CharInflation
                        tmpGlyphOutlineCopy.ConvertPath_InflateLikeBalloon m_CharInflation, P2_LJ_Round, P2_LC_Round
                        
                    End If
                    
                    'Finally, jitter applies some randomness to the character's position in the x and/or y directions.
                    ' (Note that at high levels, this *can* move glyphs outside the bounding box - but the user has
                    ' access to padding parameters, so we leave it to them to solve this for now.)
                    With m_Glyphs(curGlyph)
                        
                        If (m_CharJitterX <> 0!) Or (m_CharJitterY <> 0!) Then
                            dstPath.AddPath tmpGlyphOutlineCopy, .finalX + .glyphOffset.du + ((cRandom.GetRandomFloat_VB * maxJitterX) - halfJitterX), .finalY - .glyphOffset.dv + ((cRandom.GetRandomFloat_VB * maxJitterY) - halfJitterY)
                        Else
                            dstPath.AddPath tmpGlyphOutlineCopy, .finalX + .glyphOffset.du, .finalY - .glyphOffset.dv
                        End If
                        
                    End With
                    
                'If no weird character effects are active, we can just render the glyph as-is
                Else
                
                    With m_Glyphs(curGlyph)
                        dstPath.AddPath m_GlyphOutlines(curGlyphOutline).glyphPath, .finalX + .glyphOffset.du, .finalY - .glyphOffset.dv
                    End With
                    
                End If
                
            End If
            
            'IMPORTANT NOTE:
            ' The following chunk of commented-out lines can be used to apply underline/strikeout behavior
            ' on a per-glyph basis. At present, PD treats an underline/strikeout rect as a line-level continuous
            ' shape.  This is done to avoid separator lines appearing if the user renders text with some kind of
            ' outline feature.
            '
            ' To enable the old behavior, uncomment the lines below.  (I'm leaving them as it may be nice to add
            ' this as a toggle-able feature in the future.
            
'            'Regardless of glyph renderability, we always apply underline and strikeout (if active, obviously).
'            If m_StyleUnderline And m_Glyphs(curGlyph).AdvanceWidth <> 0 Then
'
'                underlineRect.Left = m_Glyphs(curGlyph).finalX
'                underlineRect.Width = m_Glyphs(curGlyph).AdvanceWidth
'
'                underlineRect.Top = m_Glyphs(curGlyph).finalY + m_TextMetrics.tmAscent - m_OutlineTextMetrics.otmsUnderscorePosition
'                underlineRect.Height = m_OutlineTextMetrics.otmsUnderscoreSize
'
'                dstPath.addRectangle_RectF underlineRect
'
'            End If
'
'            If m_StyleStrikeout And m_Glyphs(curGlyph).AdvanceWidth <> 0 Then
'
'                underlineRect.Left = m_Glyphs(curGlyph).finalX
'                underlineRect.Width = m_Glyphs(curGlyph).AdvanceWidth
'
'                underlineRect.Top = m_Glyphs(curGlyph).finalY + m_TextMetrics.tmAscent - m_OutlineTextMetrics.otmsStrikeoutPosition
'                underlineRect.Height = m_OutlineTextMetrics.otmsStrikeoutSize
'
'                dstPath.addRectangle_RectF underlineRect
'
'            End If
            
        'Render the next glyph!
        Next i
        
        'With all glyphs assembled, apply underline/strikeout if any
        If m_StyleUnderline Or m_StyleStrikeout Then dstPath.AddPath additionalStylesRect
        
        'If stretch-to-fit mode is active, manually squeeze the final path into the passed bounding box
        If (stretchToFit = stf_Box) Then
            
            Dim cTransform As pd2DTransform
            Set cTransform = New pd2DTransform
            
            Dim unstretchedBounds As RectF
            unstretchedBounds = dstPath.GetPathBoundariesF
            
            'Failsafe check for DBZ
            If (boundingRect.Width > 0!) And (boundingRect.Height > 0!) Then
                
                Dim oldWidth As Single, oldHeight As Single
                oldWidth = unstretchedBounds.Width
                If (oldWidth <= 0!) Then oldWidth = boundingRect.Width
                oldHeight = unstretchedBounds.Height
                If (oldHeight <= 0!) Then oldHeight = boundingRect.Height
                
                'To account for center- and right-alignment, we first need to translate the existing path
                ' to the actual bounding box left, *then* apply the scaling transform.
                cTransform.ApplyTranslation boundingRect.Left - unstretchedBounds.Left, boundingRect.Top - unstretchedBounds.Top
                cTransform.ApplyScaling boundingRect.Width / oldWidth, boundingRect.Height / oldHeight, boundingRect.Left, boundingRect.Top
                dstPath.ApplyTransformation cTransform
                
            End If
                
        End If
        
    Else
        Debug.Print "No glyphs in the current path; pdGlyphCollection.AssembleCompositePath() exited without doing anything..."
    End If
        
    AssembleCompositePath = True

End Function

'Fill the official xOffset and yOffset parameters of every glyph.  Note that all the initial positioning data comes
' from Uniscribe; this function is primarily responsible for line breaks, and any PD-specific positioning changes
' (e.g. modified line or character spacing).
Private Function CalculateGlyphPositions(ByRef boundingRect As RectF, ByVal horizontalAlignment As GP_StringAlignment, ByVal verticalAlignment As GP_StringAlignment, Optional ByVal lineWrapMode As PD_TextWordwrap = tww_AutoWord, Optional ByRef dstStylePath As pd2DPath, Optional ByVal stretchToFit As PD_TextStretchToFit = stf_None, Optional ByVal justifyLastLine As GP_StringAlignment = StringAlignmentNear) As Boolean
    
    Dim i As Long, j As Long
    Dim curGlyph As Long, glyphCheck As Long
    Dim timeToBreakLine As Boolean
    
    Dim xOffset As Single, yOffset As Single
    yOffset = boundingRect.Top
    xOffset = boundingRect.Left
    
    'Pre-calculate right and bottom bounds, since we'll be using those a lot throughout this function
    Dim boundingRectRight As Single, boundingRectBottom As Single
    boundingRectRight = boundingRect.Left + boundingRect.Width
    boundingRectBottom = boundingRect.Top + boundingRect.Height
    
    'Debug.Print "Bounds for this text region: (" & boundingRect.Left & ", " & boundingRect.Top & ") - (" & boundingRectRight & ", " & boundingRectBottom & ")"
    
    Dim currentLine As Long, linebreakImpossible As Boolean, lastGlyphLineVerified As Long
    currentLine = 0
    linebreakImpossible = False
    
    'If stretch-to-fit mode is active, we cannot auto-calculate linebreaks.  The only line breaks we
    ' respect in stretch-to-fit mode are hard-coded linebreaks from the original string.
    If (stretchToFit <> stf_None) Then lineWrapMode = tww_Manual
    
    'Prior to processing each glyph, cache some unchanging values (like line increments).
    
    'Default lineHeight is just the text metric height (not counting external leading;
    ' this method is identical to GDI, but differs from GDI+)
    Dim lineHeight As Single
    lineHeight = m_TextMetrics.tmHeight
    
    'We further adjust line height by the user's specified line spacing modifier
    ' (where 0.0 means use default line height)
    lineHeight = lineHeight + (m_LineSpacing / 100!) * lineHeight
    
    'Previously, the lineHeight calculation included the m_TextMetrics.tmExternalLeading value.
    ' I no longer use this, as it causes some fonts to greatly over-extend each line.
    
    'Average char width is used to increase/decrease char spacing (if the user has variable character spacing applied)
    Dim customCharOffset As Single
    customCharOffset = CSng(m_TextMetrics.tmAveCharWidth) * m_CharSpacing
    
    'Reset various glyph metrics
    For i = 0 To UBound(m_Glyphs)
        With m_Glyphs(i)
            .isFirstGlyphOnLine = False
            .isLastGlyphOnLine = False
            .lineID = 0
        End With
    Next i
    
    'Mark the first glyph as first on its line; we use this indicator to prevent infinite loops if wordwrap
    ' is activated, but we can't fit an entire word on a line.  When this happens, we have no choice but to
    ' clip the word.  (This is the same behavior as modern word processors, to my knowledge.)
    m_Glyphs(0).isFirstGlyphOnLine = True
    
    'Reset the line widths array; new line widths are about to be calculated!
    For i = 0 To UBound(m_LineWidths)
        m_LineWidths(i) = 0
    Next i
    
    'For performance reasons, we calculate all initial positioning in a single pass.  I had originally intended
    ' to make two passes - one to determine line break positions, and a second to perform actual (x, y) positioning.
    ' This would make for cleaner code, but it also wastes time, so I've condensed everything into a single,
    ' multipurpose pass.
    '
    'Note that a second pass is still required for repositioning due to non-left or non-top alignments;
    ' this is a necessary evil because we can't apply specialized vertical positioning until we've calculated *all*
    ' linebreaks, and we can't calculate linebreaks until we calculate all x-positioning.  However, we do what we can
    ' to improve performance by caching calculated line widths as we go.
    Dim curGlyphWidth As Single, lineWidthAdded As Single
    curGlyphWidth = 0
    lineWidthAdded = xOffset
    
    'Start iterating through the source string, one character at a time.
    curGlyph = 0
    
    Do
        
        'How we determine line breaks depends on the current wordwrap modes.
        
        'In non-break mode, all line breaks (even manual ones) are forcibly ignored
        If (lineWrapMode = tww_None) Then
            timeToBreakLine = False
            
        'In all other modes, hard breaks (e.g. \n) are respected
        Else
            
            'If the current glyph is a line break glyph, use it preferentially over any other cues
            If m_Glyphs(curGlyph).isHardLineBreak Then
                timeToBreakLine = True
            
            'If this is not a line break glyph, handle it according to the user's settings
            Else
                
                'By default, assume a linebreak is *not* required
                timeToBreakLine = False
                
                'Regardless of break mode, start by seeing if this character extends past the end of the line.
                ' Note that this character's C-overhang is included, to avoid clipping on italicized characters.
                If (xOffset + m_Glyphs(curGlyph).advanceWidth + (-1 * m_Glyphs(curGlyph).abcWidth.abcC) >= boundingRectRight) Then
                
                    'This character doesn't fit.  Further handling is determined by wrap mode.
                    
                    'In character mode, lines are broken as soon as a character extends past a line,
                    ' so simply mark the break and move on.
                    If (lineWrapMode = tww_AutoCharacter) Then
                        timeToBreakLine = True
                    
                    'In word mode, language-specific word-wrap rules are used.  This is complicated, because we
                    ' must forcibly retract the character pointer to the last viable breaking position.  If that
                    ' happens to be the first character on this line, it means this word doesn't physically fit,
                    ' so we have no choice but to break in the middle of the word.
                    ElseIf (lineWrapMode = tww_AutoWord) Then
                        
                        'One way or another, we're going to apply a linebreak
                        timeToBreakLine = True
                        
                        'If a legit linebreak isn't possible (e.g. a single word extends from the start of the line past the
                        ' end of the line), we'll forcibly break the word in two.
                        linebreakImpossible = False
                        
                        'If this glyph is breakable, awesome!  Break immediately.
                        If (m_Glyphs(curGlyph).isSoftBreak Or m_Glyphs(curGlyph).isWhiteSpace) Then
                            
                            'No extra code required
                            
                        'If this glyph is *not* breakable (meaning it sits in the middle of a word), loop backward
                        ' until we find a glyph that *is* breakable.  Point the glyph pointer at that glyph;
                        ' everything past it must be re-positioned on a new line.
                        '
                        '(In the future, this is where we could add hyphenation support.)
                        Else
                            
                            'Failsafe check for a line break being required on the initial glyph; this should never happen,
                            ' but if it does, skip linebreaking as there's no point.
                            If (curGlyph = 0) Then
                                timeToBreakLine = False
                            Else
                                
                                'Note that we have not checked any glyphs surrounding this line break.
                                ' (We need to know this when testing for "impossible to break" lines, below.)
                                lastGlyphLineVerified = -1
                                
                                'Loop backward until we find a viable breaking position.
                                For j = curGlyph - 1 To 0 Step -1
                                
                                    'To avoid infinite loops, it's important to check failure states first.
                                    
                                    'If we encounter a hard line break, this word is unbreakable.  Mark failure and exit.
                                    If m_Glyphs(j).isHardLineBreak Then
                                        linebreakImpossible = True
                                        Exit For
                                    End If
                                    
                                    'If we encounter the first character in this line (or the last character on a
                                    ' previous line, as a failsafe), this word is unbreakable.  Mark failure and exit.
                                    If m_Glyphs(j).isFirstGlyphOnLine Then
                                        linebreakImpossible = True
                                        Exit For
                                    ElseIf m_Glyphs(j).isLastGlyphOnLine Then
                                        linebreakImpossible = True
                                        Exit For
                                    End If
                                    
                                    'If we're still here, it means this character is a viable candidate for a valid break.
                                    
                                    'Remove its width from the running line width counter.  (Tracking line width here is
                                    ' convenient for a later step - aligning and/or justifying lines, which require line width
                                    ' as part of their calculations.)
                                    If ((m_Glyphs(j).advanceWidth + customCharOffset) > 0) Then
                                        m_LineWidths(currentLine) = m_LineWidths(currentLine) - (m_Glyphs(j).advanceWidth + customCharOffset)
                                    Else
                                        m_LineWidths(currentLine) = m_LineWidths(currentLine) - m_Glyphs(j).advanceWidth
                                    End If
                                    
                                    'Because this glyph has been removed from the running line width, note that it has been
                                    ' "checked". (We need this index in case this line is un-breakable; when that happens,
                                    ' we need to re-add all removed line widths to the line in question.)
                                    lastGlyphLineVerified = j
                                    
                                    'If this is a valid softbreak position, mark it and exit immediately
                                    If (m_Glyphs(j).isSoftBreak Or m_Glyphs(j).isWhiteSpace) Then
                                        curGlyph = j
                                        Exit For
                                    End If
                                    
                                Next j
                                
                                'Note that j reaching 0 is okay; this just means we reached the start of the string,
                                ' but didn't find a valid breakpoint.
                                If (j = 0) Then linebreakImpossible = True
                                
                                'If a line break isn't possible, we have no choice but to break on this character.
                                ' (Again, in the future this could consider hyphenation breakpoints as well.)
                                '
                                'Leave the curGlyph pointer where it is, but (and I know this is weird) loop back through
                                ' all the characters we just removed from our running line width.  Since they are now "stuck"
                                ' on the current line, we have no choice but to add their widths back into the current
                                ' line's width.
                                If linebreakImpossible Then
                                
                                    'If one or more glyph widths were removed from this line, add them back in now.
                                    If (lastGlyphLineVerified >= 0) Then
                                        
                                        'Because the current glyph is the invalid one (e.g. the one that will be
                                        ' forcibly "shoved" onto the next line), we don't want to include it in our
                                        ' line width calculations.
                                        If (lastGlyphLineVerified < curGlyph) Then
                                            
                                            For j = lastGlyphLineVerified To curGlyph - 1
                                            
                                                'This glyph is stuck on the current line.
                                                ' Return its width to the running line width counter.
                                                If ((m_Glyphs(j).advanceWidth + customCharOffset) > 0) Then
                                                    m_LineWidths(currentLine) = m_LineWidths(currentLine) + (m_Glyphs(j).advanceWidth + customCharOffset)
                                                Else
                                                    m_LineWidths(currentLine) = m_LineWidths(currentLine) + m_Glyphs(j).advanceWidth
                                                End If
                                            
                                            Next j
                                            
                                        End If
                                        
                                    End If
                                
                                End If
                                
                            'End failsafe check (If/Else curGlyph = 0 and we're being forced to break)
                            End If
                        
                        'End If/Then for being able to break on the exact current character, or being forced to search for a new breakpoint.
                        End If
                            
                    'End If/Then for current line break mode
                    End If
                
                'This character fits on the line.  Carry on!
                Else
                    timeToBreakLine = False
                End If
                    
            'End If/Else for hard line breaks
            End If
            
        'End If/Else for lineBreakMode = no breaking
        End If
        
        'If this glyph (or in the case of wordwrap, some previous glyph which the counter now points to)
        ' is marked for line breaking, proceed to break the line now.
        If timeToBreakLine Then
        
            'Increment the current line counter
            currentLine = currentLine + 1
            If (currentLine > UBound(m_LineWidths)) Then ReDim Preserve m_LineWidths(0 To currentLine * 2 - 1) As Single
            
            'Increment the running y-offset by the preset lineheight, and reset the x-offset to 0
            ' (or more accurately, the current left boundary, as specified by our caller)
            yOffset = yOffset + lineHeight
            xOffset = boundingRect.Left
            m_LineWidths(currentLine) = 0
            
            'We now need to mark the last character of the previous line, and the first character of this line.
            ' (This is important for future wordbreak calculations.)  Note that the character in question may
            ' or may not be the current glyph, *if* the current glyph is a whitespace character.
            
            'First, deal with hard linebreaks.  They are a special circumstance, and if one is used, we ALWAYS
            ' mark it as the last glyph of the preceding paragraph.  (This is necessary if the first glyph in the
            ' string is a linebreak, for example.)
            If m_Glyphs(curGlyph).isHardLineBreak Then
                m_Glyphs(curGlyph).isLastGlyphOnLine = True
                m_Glyphs(curGlyph).lineID = currentLine - 1   'while we're here, ensure the line ID is correct
                
            'If we are not breaking on a hard line break, we mark the *previous* non-whitespace character
            ' as the last one on the line.
            Else
                
                If (curGlyph <> 0) Then
                    
                    'We now need to search backward for the last non-whitespace glyph.
                    For glyphCheck = curGlyph - 1 To 0 Step -1
                        
                        'If this glyph is non-whitespace, it's the one we want.  Mark it and exit the loop.
                        If (Not m_Glyphs(glyphCheck).isWhiteSpace) Then
                            m_Glyphs(glyphCheck).isLastGlyphOnLine = True
                            Exit For
                        
                        'This is a whitespace glyph.
                        Else
                        
                            'Remove this character's width from our running line width.
                            ' We *do not* want to factor end-of-line whitespace into positioning calculations.
                            If ((m_Glyphs(glyphCheck).advanceWidth + customCharOffset) > 0) Then
                                m_LineWidths(currentLine - 1) = m_LineWidths(currentLine - 1) - (m_Glyphs(glyphCheck).advanceWidth + customCharOffset)
                            Else
                                m_LineWidths(currentLine - 1) = m_LineWidths(currentLine - 1) - m_Glyphs(glyphCheck).advanceWidth
                            End If
                        
                        End If
                        
                    Next glyphCheck
                    
                    'If glyphCheck made it all the way to 0, the previous line consists of nothing but whitespace characters.
                    ' In the absence of anything better, mark the last one as the last glyph on the line.
                    If (glyphCheck = 0) Then
                        m_Glyphs(curGlyph - 1).isLastGlyphOnLine = True
                    
                    'glyphCheck is non-zero, meaning the previous line ends with a non-white space character.
                    Else
                    
                        'Add the "C" overhang from this character to the line width.  This is important for getting right-
                        ' and center-aligned text *perfect*.
                        m_LineWidths(currentLine - 1) = m_LineWidths(currentLine - 1) - m_Glyphs(glyphCheck).abcWidth.abcC
                    
                    End If
                
                'If this is the first glyph in the string (seriously?), note that it is both the first AND last glyph
                ' on its line.
                Else
                    m_Glyphs(0).isLastGlyphOnLine = True
                End If
            
            '/end finding+marking the last character of the previous line
            End If
            
            'Next, we need to mark the first glyph of the next line.  This mark is important for subsequent
            ' line handling, particularly justified line alignment.
            
            'We do this by advancing the glyph pointer beyond its current point until we encounter a non-whitespace,
            ' non-hard-linebreak glyph.  (This ensures that trailing white space behaves correctly.)
            
            'Note that we must pre-check for the glyph pointer being valid; it may point beyond the end of the array
            ' if the last character in a string is a hard line-break.
            If (curGlyph < m_NumOfGlyphs) Then
                
                'If the current glyph is a non-whitespace glyph, mark it as the start of the next line.
                If (Not m_Glyphs(curGlyph).isWhiteSpace) Then
                    m_Glyphs(curGlyph).isFirstGlyphOnLine = True
                
                'If this glyph *is* a whitespace glyph, we have extra work to do.
                Else
                    
                    'Failsafe checks for "glyph counter beyond end of string" are always advised
                    Dim incrementGlyphPointer As Boolean
                    incrementGlyphPointer = (curGlyph < m_NumOfGlyphs)
                    
                    'Advance the glyph pointer until we reach either...
                    ' 1) a non-whitespace glyph, or
                    ' 2) a hard line-break
                    Do While incrementGlyphPointer
                        
                        'The current character is not whitespace (so it's a valid "first of line" character).
                        If (Not m_Glyphs(curGlyph).isWhiteSpace) Then Exit Do
                        
                        'The current character is a hard linebreak.  Advance the glyph pointer to the next glyph
                        ' (whatever it is) because it must be the start of the next line.
                        If m_Glyphs(curGlyph).isHardLineBreak Then
                            curGlyph = curGlyph + 1
                            If (curGlyph <= m_NumOfGlyphs - 1) Then m_Glyphs(curGlyph).isFirstGlyphOnLine = True
                            GoTo SkipToNextGlyph
                        End If
                        
                        'If we're still here, this is a whitespace glyph.  Make sure the line ID is correct
                        ' (treating this glyph as if it belongs to the *previous* line), then increment curGlyph
                        ' and continue looking for a non-whitespace glyph or a new linebreak char.
                        m_Glyphs(curGlyph).lineID = currentLine - 1
                        curGlyph = curGlyph + 1
                        incrementGlyphPointer = (curGlyph < m_NumOfGlyphs)
                        
                    Loop
                    
                    'We now have one of two cases:
                    ' 1) curGlyph points beyond the end of the glyph array, because no acceptable starting
                    '    line character was found.
                    ' 2) curGlyph points at the first glyph of the new line.
                    
                    'Only the second case needs to be dealt with.
                    If (curGlyph < m_NumOfGlyphs) Then m_Glyphs(curGlyph).isFirstGlyphOnLine = True
                    
                End If
                    
            End If
            
            'curGlyph now points at one of two things:
            ' 1) The first glyph on a *new* line
            ' 2) Somewhere past the end of the character array
            '
            'Only the first case needs to be dealt with, via the "If curGlyph < m_NumOfGlyphs" statement below
            
        'End of "If timeToBreakLine"...
        End If
        
        'After all that work, we can finally place this damn glyph
        If (curGlyph < m_NumOfGlyphs) Then
        
            'With line breaks now accounted for, glyph positioning is pretty darn simple.
            
            'Y-positioning is universal, but X-positioning changes for characters who are first in line.
            If m_Glyphs(curGlyph).isFirstGlyphOnLine Then
                
                'For initial glyphs on a line, we increment the x-offset artificially, using their A overhang.
                ' This is important for avoiding clipping on synthesized italic fonts (e.g. Times New Roman at large sizes)
                xOffset = boundingRect.Left + 1 + (-1 * m_Glyphs(curGlyph).abcWidth.abcA)
                
            End If
            
            'This character can now be positioned!
            With m_Glyphs(curGlyph)
                .finalX = xOffset
                .finalY = yOffset
                .lineID = currentLine
            End With
            
            'Increment the running x-offset with this character's width
            curGlyphWidth = m_Glyphs(curGlyph).advanceWidth
            
            'We also apply the user's custom character spacing parameter, if applicable.
            ' (If applying it would result in a negative offset relative to the previous character, we ignore it.)
            If (curGlyphWidth + customCharOffset > 0) Then
                xOffset = xOffset + curGlyphWidth + customCharOffset
            End If
            
            'Keep a running tally of the current line widths; this makes our second alignment pass much faster.
            If (currentLine > UBound(m_LineWidths)) Then ReDim Preserve m_LineWidths(0 To currentLine * 2 - 1) As Single
            m_LineWidths(currentLine) = xOffset
            
            'Debug information can be helpful when dealing with such complicated measurement systems:
            'Debug.Print "Glyph: " & curGlyph & ", line: " & currentLine & ", linewidth: " & m_LineWidths(currentLine) & ", offset: " & xOffset
                
            'Advance to the next character
            curGlyph = curGlyph + 1
            
        End If

SkipToNextGlyph:
    Loop While (curGlyph < m_NumOfGlyphs)
    
    'After all glyphs are placed, find the last character in the string, and make sure it's marked as
    ' the last one on its line.
    If (m_NumOfGlyphs > 0) Then
    
        m_Glyphs(m_NumOfGlyphs - 1).isLastGlyphOnLine = True
        
        'We also want to add the C measurement of this glyph to the running line width of the final line,
        ' since that's normally handled by the "move to next line" code, above.
        m_LineWidths(currentLine) = m_LineWidths(currentLine) - m_Glyphs(m_NumOfGlyphs - 1).abcWidth.abcC
        
    Else
        m_Glyphs(0).isLastGlyphOnLine = True
    End If
    
    'At this point, each glyph is now correctly positioned according to a default left/top alignment.
    ' We have also marked the first and last glyph on every line, and assigned each glyph a line ID.
        
    'As a convenience to subsequent functions, we are now going to mark all-whitespace lines as being zero-width.
    ' This solves some complex problems when attempting to place underline and strikeout styles, as whitespace-only
    ' lines should *not* have underline/strikeout applied.
    Dim curLine As Long, numOfLines As Long
    numOfLines = currentLine + 1
    
    Dim lineHasNonWhitespace As Boolean, lastGlyphChecked As Long
    lastGlyphChecked = 0
    
    For curLine = 0 To numOfLines - 1
        
        lineHasNonWhitespace = False
        
        'Search for glyphs that 1) sit on this line, and 2) are not whitespace
        For i = lastGlyphChecked To m_NumOfGlyphs - 1
        
            'If a non-whitespace char is found on this line, note it and exit immediately
            If ((m_Glyphs(i).lineID = curLine) And (Not m_Glyphs(i).isWhiteSpace)) Then
                lineHasNonWhitespace = True
                lastGlyphChecked = i
                Exit For
            End If
                        
        Next i
        
        'If no whitespace chars were found on this line, mark it as zero-width
        If (Not lineHasNonWhitespace) Then m_LineWidths(curLine) = 0
    
    Next curLine
    
    'If the text alignment is anything other than left/top, we now need to apply additional passes to
    ' the glyph collection. The line information assembled by our previous runs is crucial here.
    Dim lineDiff As Single
    Dim hLineOffsets() As Single
    ReDim hLineOffsets(0 To UBound(m_LineWidths)) As Single
    
    'Start with horizontal alignment
    If (horizontalAlignment <> StringAlignmentNear) Then
        
        'Justified text requires a lot of special handling, alas...
        If (horizontalAlignment = StringAlignmentJustify) Then
            
            'For justified text to work, we need to do a few different things.
            ' (Note that the current algorithm works on a per-line basis.  A more sophisticated algorithm
            ' might rework neighboring lines to "push" a word up or down - a la the classic Knuth-Plass algorithm -
            ' or we could hyphenate words to make the justified text look prettier.  These more advanced approaches
            ' are often language-specific which is why I haven't tackled them... yet.)
            
            'Start by iterating lines.
            lastGlyphChecked = 0
            For curLine = 0 To numOfLines - 1
                
                'Skip whitespace-only lines (which were explicitly marked as 0-length in a previous step).
                ' Note that we may also skip the last line in multi-line paragraphs, per the user's
                ' "last line justify" setting, but we still need to iterate all lines here to look for
                ' last lines of multi-line paragraphs that are *not* also the last line in the layer.
                ' (This is possible if a single text layer contains multiple paragraphs.)
                Dim justifyThisLine As Boolean
                justifyThisLine = (m_LineWidths(curLine) > 0)
                
                Dim lineEndsInHardBreak As Boolean
                lineEndsInHardBreak = False
                
                If justifyThisLine Then
                    
                    'If possible, we want to insert extra space only where whitespace characters already exist.
                    ' To do that, we need to know how many whitespace characters we have to work with on this line.
                    Dim numWhiteSpaceChars As Long
                    numWhiteSpaceChars = 0
                    
                    Dim idxFirstChar As Long, idxLastChar As Long
                    idxFirstChar = -1: idxLastChar = -1
                    
                    'Search for glyphs that 1) sit on the current line, and 2) are not whitespace
                    For i = lastGlyphChecked To m_NumOfGlyphs - 1
                    
                        'Line ID must match the current line we are attempting to justify
                        If (m_Glyphs(i).lineID = curLine) Then
                            
                            'Note first/last glyph indices for this line, and deliberately omit them from justification
                            ' (by marking them as zero-width)
                            If m_Glyphs(i).isFirstGlyphOnLine Then
                                idxFirstChar = i
                                If m_Glyphs(i).isWhiteSpace Then m_Glyphs(i).isZeroWidth = True
                            End If
                            
                            If m_Glyphs(i).isLastGlyphOnLine Then
                                lineEndsInHardBreak = m_Glyphs(i).isHardLineBreak
                                idxLastChar = i
                                If m_Glyphs(i).isWhiteSpace Then m_Glyphs(i).isZeroWidth = True
                            End If
                            
                            'If this character is a non-zero-width whitespace glyph, note that it's a
                            ' valid whitespace marker.
                            If (m_Glyphs(i).isWhiteSpace) And (Not m_Glyphs(i).isZeroWidth) Then
                                numWhiteSpaceChars = numWhiteSpaceChars + 1
                            End If
                            
                            'If this is the final valid character in this line, increment the glyph pointer
                            ' until we reach the first glyph on the *next* line.  (This fixes weird behavior
                            ' on multi-whitespace-chars at the end of a line, like a double-space after a period.)
                            If (idxLastChar = i) Then
                                Do While m_Glyphs(i).lineID = curLine
                                    i = i + 1
                                    If (i >= m_NumOfGlyphs) Then Exit For
                                Loop
                                lastGlyphChecked = i
                                Exit For
                            End If
                        
                        'Blank lines in a row (e.g. LF + LF) can trigger this branch.  When this happens,
                        ' we want to manually advance the glyph pointer to the first glyph of the *next* line.
                        Else
                            
                            If (m_Glyphs(i).lineID > curLine) Then
                                PDDebug.LogAction "WARNING: pdGlyphCollection.CalculateGlyphPositions bad parse."
                            
                            'curLine must be lower than the expected line
                            Else
                                
                                'Manually advance the glyph pointer to the start of the next line
                                lastGlyphChecked = i
                                Do
                                    lastGlyphChecked = lastGlyphChecked + 1
                                    If (lastGlyphChecked > m_NumOfGlyphs - 1) Then Exit Do
                                Loop While (m_Glyphs(lastGlyphChecked).lineID > curLine)
                                
                                'i is a loop counter and it will be auto-incremented before touching the next glyph,
                                ' so retreat it by 1 so the next iteration starts at "lastGlyphChecked".
                                i = lastGlyphChecked - 1
                                
                            End If
                        
                        '/end m_Glyphs(i).lineID = curLine
                        End If
                        
                    Next i
                    
                    'Before proceeding, we want to forcibly strip whitespace glyphs from the end of each "line".
                    ' (On a softbreak, whitespace characters are allowed to "drift beyond the margin", but we don't want
                    ' to consider those whitespace chars as valid targets for inserting justifying space.)
                    Do While (idxLastChar > idxFirstChar) And (idxLastChar >= 0)
                        If m_Glyphs(idxLastChar).isWhiteSpace Then
                            m_Glyphs(idxLastChar).isZeroWidth = True
                            idxLastChar = idxLastChar - 1
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    'We have now "trimmed" the last character indices to point only at the last non-whitespace-char
                    ' on this line.  (Note that preceding whitespace chars are fine - this allows the user to insert
                    ' e.g. spaces on the first line of a paragraph, and justification will still work correctly.)
                    
                    'Next, check a couple weird failure states for odd text arrangements.
                    
                    'Single-character lines cannot be justified
                    If (idxFirstChar >= idxLastChar + 1) Then GoTo NextLineJustify
                    
                    'Failsafe only for bad line parsing (should never trigger)
                    If (idxFirstChar < 0) Or (idxLastChar < 0) Then GoTo NextLineJustify
                    
                    'Figure out how much space we need to "insert" across the current line to make it justified.
                    lineDiff = boundingRectRight - m_LineWidths(curLine)
                    
                    'Next, look for trailing lines in a paragraph.  Trailing lines in a paragraph may need to be
                    ' dealt with specially, depending on the user's current "last line justify" setting.
                    Dim isTrailingLine As Boolean
                    isTrailingLine = False
                    
                    'Trailing lines are either the last line in the paragraph (by default) or any inter-paragraph line
                    ' that terminates in a hard linebreak.  (I'm proud that PD handles that second case correctly -
                    ' other photo editors, like GIMP, do not!)
                    If (numOfLines > 1) Then
                        isTrailingLine = (curLine = numOfLines - 1) Or lineEndsInHardBreak
                    End If
                    
                    'Note that we only need to handle trailing lines specially if the user has *not* specified
                    ' last-line justification...
                    If isTrailingLine And (justifyLastLine <> StringAlignmentJustify) Then
                        
                        'The user wants the last line of each paragraph justified left/center/right.
                        
                        'Left-alignment doesn't actually require any work (it's applied by default).
                        If (justifyLastLine <> StringAlignmentNear) Then
                        
                            'Calculate a new per-glyph offset required for the target alignment.
                            If (justifyLastLine = StringAlignmentCenter) Then lineDiff = lineDiff * 0.5!
                            
                            'Apply the offset to each glyph in this line.
                            For j = idxFirstChar To idxLastChar
                                If (Not m_Glyphs(j).isZeroWidth) Then
                                    m_Glyphs(j).finalX = m_Glyphs(j).finalX + lineDiff
                                End If
                            Next j
                        
                        End If
                        
                        'Continue with the next line (instead of allowing the justify algorithm to continue)
                        GoTo NextLineJustify
                        
                    End If
                    
                    'We now need to split based on the number of whitespace chars in the current line.  This works if
                    ' we have at least *one* whitespace char (that is not leading or trailing), but if we have zero,
                    ' we instead need to use inter-character spacing.
                    '
                    '(An ideal implementation would likely use a mix of whitespace and inter-character spacing to
                    ' achieve "prettier" results, but I have not tackled this... yet.)
                    Dim perCharIncrement As Single, accumulatedSpacing As Single
                    accumulatedSpacing = 0!
                    
                    If (numWhiteSpaceChars > 0) Then
                        
                        'Iterate through all glyphs, adding spacing as we go.
                        perCharIncrement = lineDiff / numWhiteSpaceChars
                        
                        For j = idxFirstChar To idxLastChar
                            
                            'Ignore zero-width chars (including zero-width whitespace)
                            If (Not m_Glyphs(j).isZeroWidth) Then
                                If m_Glyphs(j).isWhiteSpace Then accumulatedSpacing = accumulatedSpacing + perCharIncrement
                                m_Glyphs(j).finalX = m_Glyphs(j).finalX + accumulatedSpacing
                            End If
                        Next j
                        
                        'Because this line was forcibly set to "full-width", set the display line-width to match.
                        ' (This allows underline and strikethrough modes to behave correctly.)
                        m_LineWidths(curLine) = boundingRectRight
                        
                    'When no whitespace chars are available, we must use inter-character spacing instead.  Yay?
                    Else
                        
                        'Some whitespace chars may still exist in the string (like the trailing one associated
                        ' with the wordbreak), but they will have been deliberately marked as zero-width in a
                        ' previous step.  We now need to count how many *actual* glyphs - e.g. non-zero-width
                        ' glyphs - exist on this line.
                        Dim numNonZeroWidth As Long
                        numNonZeroWidth = 0
                        For j = idxFirstChar To idxLastChar
                            If (Not m_Glyphs(j).isZeroWidth) Then numNonZeroWidth = numNonZeroWidth + 1
                        Next j
                        
                        'Ensure we have at least two characters to work with
                        If (idxLastChar > idxFirstChar) Then
                            
                            'Every character will receive a fixed addition to its x-position.
                            perCharIncrement = lineDiff / (numNonZeroWidth - 1)
                            
                            For j = idxFirstChar + 1 To idxLastChar
                                If (Not m_Glyphs(j).isZeroWidth) Then
                                    accumulatedSpacing = accumulatedSpacing + perCharIncrement
                                    m_Glyphs(j).finalX = m_Glyphs(j).finalX + accumulatedSpacing
                                End If
                            Next j
                            
                        End If
                        
                        'Because this line was forcibly set to "full-width", set the display line-width to match.
                        ' (This is necessary for underline and strikethrough modes, to ensure the line stretches
                        ' across the entirety of this line.)
                        m_LineWidths(curLine) = boundingRectRight
                        
                    End If
                    
                End If

NextLineJustify:
            Next curLine
            
        'Right/center alignment is much easier
        Else
            
            'Because we kept a running tally of each line's width during the processing stages, it is easy
            ' to calculate offsets now.
            For curLine = 0 To numOfLines - 1
            
                'For each line, replace its calculated line width with the difference between the bounding rect
                ' and the actual size.
                lineDiff = boundingRectRight - m_LineWidths(curLine)
                
                If (horizontalAlignment = StringAlignmentCenter) Then
                    hLineOffsets(curLine) = lineDiff * 0.5!
                Else
                    hLineOffsets(curLine) = lineDiff
                End If
            
            Next curLine
            
            'Each line width tracker now contains the offset required to make alignment work.  Apply it to each glyph in turn.
            For curGlyph = 0 To m_NumOfGlyphs - 1
                With m_Glyphs(curGlyph)
                    .finalX = .finalX + hLineOffsets(.lineID)
                End With
            Next curGlyph
            
        End If
            
    End If
    
    'Next, vertical alignment
    If (verticalAlignment <> StringAlignmentNear) Then
        
        'Because line heights are uniform, this process is much simpler, as we don't need to custom-measure anything.
        '
        'Note, however, that we must modify the line size a bit uniquely.  The first line of text is always displayed
        ' using the raw, untouched character height (m_TextMetrics.tmHeight).  Subsequent lines are offset by
        ' *our internal line spacing calculation*, which is currently the default text height multiplied by any
        ' any user-supplied modifier(s).
        '
        'This calculation ensures that a single line of text is still positioned correctly, as the user's line-spacing
        ' modification only comes into play if *multiple* lines of text are present.
        lineDiff = boundingRectBottom - (((numOfLines - 1) * lineHeight) + (m_TextMetrics.tmHeight))
        
        If (verticalAlignment = StringAlignmentCenter) Then lineDiff = lineDiff * 0.5!
        
        'Apply the vertical offset to each glyph in turn.
        For curGlyph = 0 To m_NumOfGlyphs - 1
            m_Glyphs(curGlyph).finalY = m_Glyphs(curGlyph).finalY + lineDiff
        Next curGlyph
        
    End If
    
    'If underline or strikeout text is set, we now need to assemble underline/strikeout-specific paths.
    '
    'Why do this here?  The final path assembly loop operates on a per-glyph basis.  This function operates
    ' on a per-line basis, which is more appropriate for underline and strikeout behavior.  (Assembling them
    ' on a per-glyph basis works alright, but it can cause gaps to appear between the underline of neighboring
    ' glyphs, for example.)
    '
    'Also, this function has manually marked the last usable (e.g. non-whitespace) glyph on each line,
    ' so we can ignore trailing whitespace when applying underline/strikeout.
    Dim underlineRect As RectF, strikeoutRect As RectF
    
    If m_StyleUnderline Then
        
        If (dstStylePath Is Nothing) Then Set dstStylePath = New pd2DPath
        
        'We now need to iterate through each line in the destination path.  For each line, we'll add an underline
        ' rectangle to the destination underline path, positioned as sized correctly according to all relevant
        ' size metrics (some of which are font-specific, e.g. underline height, while some are line-specific,
        ' e.g. the length of each line not counting trailing whitespace chars).
        curGlyph = 0
        
        Do
            
            'We only want to deal with the first character on each line, as it contains sufficient positioning data
            ' for us to infer the rest of the line's contents.
            If (m_Glyphs(curGlyph).isFirstGlyphOnLine And (m_LineWidths(m_Glyphs(curGlyph).lineID) > 0)) Then
                
                'Use this glyph's position as the starting point for this underline
                underlineRect.Left = m_Glyphs(curGlyph).finalX
                
                'Width of the underline is the *full line width* (not counting trailing whitespace chars)
                underlineRect.Width = m_LineWidths(m_Glyphs(curGlyph).lineID)
        
                'Top position and height are easy, as they're primarily based on font-level metrics, not glyph-level metrics
                underlineRect.Top = m_Glyphs(curGlyph).finalY + m_TextMetrics.tmAscent - m_OutlineTextMetrics.otmsUnderscorePosition
                underlineRect.Height = m_OutlineTextMetrics.otmsUnderscoreSize
                
                'Add the completed rect to the composite "style" path
                dstStylePath.AddRectangle_RectF underlineRect
                
            End If
            
            curGlyph = curGlyph + 1
            
        Loop While (curGlyph < m_NumOfGlyphs)
        
    End If
    
    'Strikeout is pretty much identical to underlining, except for positioning and sizing of the line
    If m_StyleStrikeout Then
        
        If (dstStylePath Is Nothing) Then Set dstStylePath = New pd2DPath
        
        curGlyph = 0
        
        Do
            
            'We only want to deal with the first character on each line, as it contains sufficient positioning data
            ' for us to infer the rest of the line's contents.
            If (m_Glyphs(curGlyph).isFirstGlyphOnLine And (m_LineWidths(m_Glyphs(curGlyph).lineID) > 0)) Then
                
                'Use this glyph's position as the starting point for this strikeout
                strikeoutRect.Left = m_Glyphs(curGlyph).finalX
                
                'Width of the strikeout is the *full line width* (not counting trailing whitespace chars)
                strikeoutRect.Width = m_LineWidths(m_Glyphs(curGlyph).lineID)
        
                'Top position and height are easy, as they're primarily based on font-level metrics, not glyph-level metrics
                strikeoutRect.Top = m_Glyphs(curGlyph).finalY + m_TextMetrics.tmAscent - m_OutlineTextMetrics.otmsStrikeoutPosition
                strikeoutRect.Height = m_OutlineTextMetrics.otmsStrikeoutSize
                
                'Add the completed rect to the composite "style" path
                dstStylePath.AddRectangle_RectF strikeoutRect
                
            End If
            
            curGlyph = curGlyph + 1
            
        Loop While (curGlyph < m_NumOfGlyphs)
        
    End If
    
    'This function doesn't currently have a fail state
    CalculateGlyphPositions = True

End Function

'Given a Long-type value - which is really a FIXED struct in disguise - parse out the integer and fractional chunks
' and convert them into a normal floating-point value.
Private Function GetSingleFromFIXED(ByVal srcFixed As Long) As Single
    
    'First, retrieve the fraction portion of the Long (which is UNSIGNED)
    Dim fracPortion As Long, integerPortion As Integer
    integerPortion = ((srcFixed And &H7FFF0000) \ &H10000) Or (&H8000 And srcFixed < 0)
    
    'Next, retrieve the value portion of the Long (which is SIGNED)
    fracPortion = (srcFixed And &HFFFF&)
    
    'Finally, return these using a standard FIXED conversion
    GetSingleFromFIXED = CSng(integerPortion) + CSng(fracPortion) / 65536!
    
End Function

'Given an index in m_Glyphs(), returning a matching glyph outline index from m_GlyphOutlines.
' (Because outline assembly is expensive, we only maintain one outline copy for each unique glyph.)
' Returns a value < 0 if the requested glyph does not exist
Private Function GetGlyphOutlineIndex(ByVal origGlyphPosition As Long) As Long

    GetGlyphOutlineIndex = -1
    
    Dim i As Long
    For i = 0 To m_NumOfGlyphOutlines - 1
        If (m_Glyphs(origGlyphPosition).glyphIndex = m_GlyphOutlines(i).glyphIndex) Then
            GetGlyphOutlineIndex = i
            Exit For
        End If
    Next i
    
End Function

'Reset the current glyph collection.
' This forces all glyphs to be recreated from scratch, so use only when *absolutely* necessary.
Friend Sub ResetGlyphCollection()
    m_NumOfGlyphs = 0
    m_NumOfGlyphOutlines = 0
    ReDim m_GlyphOutlines(0 To INITIAL_GLYPH_COLLECTION_SIZE - 1) As pdGlyphOutline
    ReDim m_Glyphs(0 To INITIAL_GLYPH_COLLECTION_SIZE - 1) As PDGlyphUniscribe
End Sub

'Return a copy of a path handle for an assembled glyph.  Returns 0 if no path data exists.  Note that 0 is valid output,
' as whitespace characters do not have glyphs (for most fonts, anyway).
Friend Function GetGlyphPathHandle(ByVal glyphIndex As Long) As Long
    If (Not m_GlyphOutlines(glyphIndex).glyphPath Is Nothing) Then
        GetGlyphPathHandle = m_GlyphOutlines(glyphIndex).glyphPath.GetHandle
    Else
        GetGlyphPathHandle = 0
    End If
End Function

Private Sub Class_Initialize()
    
    m_GDIFont = 0
    
    'Create an identity matrix and cache it at class level.
    ' Note that we deliberately invert the y-value as glyphs are returned in a coordinate system relative
    ' to their baseline, whereas screen coordinates are relative to the top-left.
    ' [1, 0]
    ' [0, 1]
    With m_IdentityMatrix
        .eM11.IntValue = 1
        .eM22.IntValue = -1
    End With
    
    'Create a temporary DIB for selecting the font into
    Set m_tmpDIB = New pdDIB
    m_tmpDIB.CreateBlank 4, 4
    
    'Create a Uniscribe interface
    Set m_Uniscribe = New pdUniscribe
    
    'Create a default line width tracking array
    ReDim m_LineWidths(0 To DEFAULT_INITIAL_LINE_COUNT - 1) As Single
    
    'Reset any custom layout modifiers
    m_LineSpacing = 0
    m_CharSpacing = 0
    m_CharOrientation = 0
    m_CharJitterX = 0
    m_CharJitterY = 0
    m_CharInflation = 0
    m_CharMirror = 0
    m_CharEffectsEnabled = False
    
End Sub

Private Sub Class_Terminate()

    'Manually release our glyph collection
    ResetGlyphCollection
    
    'Manually release our Uniscribe interface
    Set m_Uniscribe = Nothing

End Sub
