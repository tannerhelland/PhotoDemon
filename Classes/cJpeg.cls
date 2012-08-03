VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJpeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'This code was originally written by John Korejwa.

'You may download the original version of this code from the following link (good as of June 2012):
' http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=50065&lngWId=1

' Many thanks to John for providing a way to natively save JPEG images in VB6



Option Explicit
Option Base 0

'Note: original comments for this class are below, but note that some modifications have been made to better
'      suit this class for use in PhotoDemon.  Please refer to PSC for the original version, and thanks to
'      John for his work.

'Class Name:   cJpeg.cls  "JPEG Encoder Class"
'Author:       John Korejwa  <korejwa@tiac.net>
'Version:      0.9 beta  [26 / November / 2003]
'
'
'Legal:
'        This class is intended for and was uploaded to www.planetsourcecode.com
'
'        This product includes JPEG compression code developed by John Korejwa.  <korejwa@tiac.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'
'
'Credits:
'        Special thanks to Barry G., a government research scientist who took an interest in my
'        steganography software and research in late 1999.  I never met Barry in person, but he
'        was kind enough to buy and mail me a book with the ISO DIS 10918-1 JPEG standard.
'
'
'Description:  This class contains code for compressing pictures, sampled via hDC, into
'              baseline .JPG files.  Please report any errors or unusual behavior to the email
'              address above.
'
'Dependencies: None
'


'JPEG Marker Constants                (Note: VB compiler does not compile unused constants)
                                      'Non-Differential Huffman Coding
Private Const SOF0    As Long = &HC0& 'Baseline DCT
Private Const SOF1    As Long = &HC1& 'Extended sequential DCT
Private Const SOF2    As Long = &HC2& 'Progressive DCT
Private Const SOF3    As Long = &HC3& 'Spatial (sequential) lossless
                                      'Differential Huffman coding
Private Const SOF5    As Long = &HC5& 'Differential sequential DCT
Private Const SOF6    As Long = &HC6& 'Differential progressive DCT
Private Const SOF7    As Long = &HC7& 'Differential spatial
                                      'Non-Differential arithmetic coding
Private Const JPG     As Long = &HC8& 'Reserved for JPEG extentions
Private Const SOF9    As Long = &HC9& 'Extended sequential DCT
Private Const SOF10   As Long = &HCA& 'Progressive DCT
Private Const SOF11   As Long = &HCB& 'Spatial (sequential) lossless
                                      'Differential arithmetic coding
Private Const SOF13   As Long = &HCD& 'Differential sequential DCT
Private Const SOF14   As Long = &HCE& 'Differential progressive DCT
Private Const SOF15   As Long = &HCF& 'Differential Spatial
                                      'Other Markers
Private Const DHT     As Long = &HC4& 'Define Huffman tables
Private Const DAC     As Long = &HCC& 'Define arithmetic coding conditioning(s)
Private Const RSTm    As Long = &HD0& 'Restart with modulo 8 count "m"
Private Const RSTm2   As Long = &HD7& 'to 'Restart with modulo 8 count "m"
Private Const SOI     As Long = &HD8& 'Start of image
Private Const EOI     As Long = &HD9& 'End of image
Private Const SOS     As Long = &HDA& 'Start of scan
Private Const DQT     As Long = &HDB& 'Define quantization table(s)
Private Const DNL     As Long = &HDC& 'Define number of lines
Private Const DRI     As Long = &HDD& 'Define restart interval
Private Const DHP     As Long = &HDE& 'Define hierarchical progression
Private Const EXP     As Long = &HDF& 'Expand reference components
Private Const APP0    As Long = &HE0& 'Reserved for application segments
Private Const APPF    As Long = &HEF& '  to Reserved for application segments
Private Const JPGn    As Long = &HF0& 'Reserved for JPEG Extentions
Private Const JPGn2   As Long = &HFD& '  to Reserved for JPEG Extentions
Private Const COM     As Long = &HFE& 'Comment
Private Const RESm    As Long = &H2&  'Reserved
Private Const RESm2   As Long = &HBF& '  to Reserved
Private Const TEM     As Long = &H1&  'For temporary use in arithmetic coding

'Consider these arrays of constants.
'They are initialized with the class and do not change.
Private QLumin(63)    As Integer 'Standard Luminance   Quantum (for 50% quality)
Private QChrom(63)    As Integer 'Standard Chrominance Quantum (for 50% quality)
Private FDCTScale(7)  As Double  'Constants for scaling FDCT Coefficients
Private ZigZag(7, 7)  As Long    'Zig Zag order of 8X8 block of samples

'API constants
Private Const BLACKONWHITE    As Long = 1 'nStretchMode constants for
Private Const COLORONCOLOR    As Long = 3 '  SetStretchBltMode() API function
Private Const HALFTONE        As Long = 4 'HALFTONE not supported in Win 95, 98, ME

Private Const BI_RGB          As Long = 0
Private Const DIB_RGB_COLORS  As Long = 0


'Variable types needed for DIBSections.
Private Type SAFEARRAYBOUND
    cElements         As Long
    lLbound           As Long
End Type
Private Type SAFEARRAY2D
    cDims             As Integer
    fFeatures         As Integer
    cbElements        As Long
    cLocks            As Long
    pvData            As Long
    Bounds(0 To 1)    As SAFEARRAYBOUND
End Type
Private Type RGBQUAD
    rgbBlue           As Byte
    rgbGreen          As Byte
    rgbRed            As Byte
    rgbReserved       As Byte
End Type
Private Type BITMAPINFOHEADER
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type
Private Type BITMAPINFO
    bmiHeader         As BITMAPINFOHEADER
    bmiColors         As RGBQUAD
End Type

'API needed for creating DIBSections for sampling and pixel access.
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection2 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long   'lplpVoid changed to ByRef
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)


'Custom variable types used for this JPEG encoding implementation
Private Type QUANTIZATIONTABLE
    Qk(63)            As Integer 'Quantization Values
    FScale(63)        As Single  'Multiplication values to scale and Quantize   FDCT output
    IScale(63)        As Single  'Multiplication values to scale and DeQuantize IDCT input
End Type
Private Type HUFFMANTABLE
    Bits(15)          As Byte    'Number of huffman codes of length i+1
    HUFFVAL(255)      As Byte    'Huffman symbol values
    EHUFSI(255)       As Long    'Huffman code size for symbol i
    EHUFCO(255)       As Long    'Huffman code      for symbol i
    MINCODE(15)       As Long    '
    MAXCODE(15)       As Long    'Largest code value for length i+1
End Type
Private Type COMPONENT
    Ci                As Long    'Component ID                       [0-255]
    Hi                As Long    'Horizontal Sampling Factor         [1-4]
    Vi                As Long    'Vertical   Sampling Factor         [1-4]
    Tqi               As Long    'Quantization Table Select          [0-3]
    Data()            As Integer 'DCT Coefficients
End Type

Private PP            As Long    'Sample Precision [8, 12]
Private YY            As Long    'Number of lines             [Image Height]
Private XX            As Long    'Number of samples per line  [Image Width]
Private Nf            As Long    'Number of components in Frame

Private HMax          As Long    'Maximum horizontal sampling frequency
Private VMax          As Long    'Maximum vertical   sampling frequency

Private m_Data()      As Byte    'JPEG File Data
Private m_Chr         As Long    'Current Character in m_Data
Private m_Ptr         As Long    'Byte index in m_Data
Private m_Bit         As Long    'Bit  index in m_Chr

Private m_Block(7, 7) As Single  'Buffer for calculating DCT

Private QTable(3)     As QUANTIZATIONTABLE  '4 Quantization Tables
Private HuffDC(3)     As HUFFMANTABLE       '4 DC Huffman Tables
Private HuffAC(3)     As HUFFMANTABLE       '4 AC Huffman Tables
Private Comp()        As COMPONENT          'Scan Components

Private m_Quality     As Long
Private m_Comment     As String



'========================================================================================
'              D I S C R E T E   C O S I N E   T R A N S F O R M A T I O N
'========================================================================================

Private Sub FDCT()
    
    Static t0   As Single   'Given an 8X8 block of discretely sampled values [m_Block(0-7, 0-7)],
    Static t1   As Single   'replace them with their (scaled) Forward Discrete Cosine Transformation values.
    Static t2   As Single   '80 (+64) multiplications and 464 additions are needed.
    Static t3   As Single   'Values are scaled on output, meaning that each of the 64 elements must be
    Static t4   As Single   'multiplied by constants for a final FDCT.  These final constants are combined
    Static t5   As Single   'with Quantization constants, so a final 64 multiplications combine the
    Static t6   As Single   'completion of the FDCT and Quantization in one step.
    Static t7   As Single
    Static t8   As Single
    Static i    As Long

    For i = 0 To 7                  'Process 1D FDCT on each row
        t0 = m_Block(i, 0) + m_Block(i, 7)
        t1 = m_Block(i, 0) - m_Block(i, 7)
        t2 = m_Block(i, 1) + m_Block(i, 6)
        t3 = m_Block(i, 1) - m_Block(i, 6)
        t4 = m_Block(i, 2) + m_Block(i, 5)
        t5 = m_Block(i, 2) - m_Block(i, 5)
        t6 = m_Block(i, 3) + m_Block(i, 4)
        t7 = m_Block(i, 3) - m_Block(i, 4)

        t7 = t7 + t5
        t8 = t0 - t6
        t6 = t6 + t0
        t0 = t2 + t4
        t2 = (t2 - t4 + t8) * 0.707106781186548   'Cos(2# * PI / 8#)
        t4 = t1 + t3
        t3 = (t3 + t5) * 0.707106781186548        'Cos(2# * PI / 8#)
        t5 = (t4 - t7) * 0.382683432365091        'Cos(3# * PI / 8#)
        t7 = t7 * 0.541196100146196 - t5          'Cos(PI / 8#) - Cos(3# * PI / 8#)
        t4 = t4 * 1.30656296487638 - t5           'Cos(PI / 8#) + Cos(3# * PI / 8#)
        t5 = t1 + t3
        t1 = t1 - t3

        m_Block(i, 0) = t6 + t0
        m_Block(i, 4) = t6 - t0
        m_Block(i, 1) = t5 + t4
        m_Block(i, 7) = t5 - t4
        m_Block(i, 2) = t8 + t2
        m_Block(i, 6) = t8 - t2
        m_Block(i, 5) = t1 + t7
        m_Block(i, 3) = t1 - t7
    Next i

    For i = 0 To 7                   'Process 1D FDCT on each column
        t0 = m_Block(0, i) + m_Block(7, i)
        t1 = m_Block(0, i) - m_Block(7, i)
        t2 = m_Block(1, i) + m_Block(6, i)
        t3 = m_Block(1, i) - m_Block(6, i)
        t4 = m_Block(2, i) + m_Block(5, i)
        t5 = m_Block(2, i) - m_Block(5, i)
        t6 = m_Block(3, i) + m_Block(4, i)
        t7 = m_Block(3, i) - m_Block(4, i)

        t7 = t7 + t5
        t8 = t0 - t6
        t6 = t6 + t0
        t0 = t2 + t4
        t2 = (t2 - t4 + t8) * 0.707106781186548   'Cos(2# * PI / 8#)
        t4 = t1 + t3
        t3 = (t3 + t5) * 0.707106781186548        'Cos(2# * PI / 8#)
        t5 = (t4 - t7) * 0.382683432365091        'Cos(3# * PI / 8#)
        t7 = t7 * 0.541196100146196 - t5          'Cos(PI / 8#) - Cos(3# * PI / 8#)
        t4 = t4 * 1.30656296487638 - t5           'Cos(PI / 8#) + Cos(3# * PI / 8#)
        t5 = t1 + t3
        t1 = t1 - t3

        m_Block(0, i) = t6 + t0
        m_Block(4, i) = t6 - t0
        m_Block(1, i) = t5 + t4
        m_Block(7, i) = t5 - t4
        m_Block(2, i) = t8 + t2
        m_Block(6, i) = t8 - t2
        m_Block(5, i) = t1 + t7
        m_Block(3, i) = t1 - t7
    Next i
End Sub




'================================================================================
'                 H U F F M A N   T A B L E   G E N E R A T I O N
'================================================================================
Private Sub OptimizeHuffman(TheHuff As HUFFMANTABLE, freq() As Long)
'Generate optimized values for BITS and HUFFVAL in a HUFFMANTABLE
'based on symbol frequency counts.  freq must be dimensioned freq(0-256)
'and contain counts of symbols 0-255.  freq is destroyed in this procedure.
    Dim i              As Long
    Dim j              As Long
    Dim k              As Long
    Dim n              As Long
    Dim V1             As Long
    Dim V2             As Long
    Dim others(256)    As Long
    Dim codesize(256)  As Long
    Dim Bits(256)      As Long
    Dim swp            As Long
    Dim swp2           As Long


    For i = 0 To 256  'Initialize others to -1, (this value terminates chain of indicies)
        others(i) = -1
    Next i
    freq(256) = 1     'Add dummy symbol to guarantee no code will be all '1' bits

   'Generate codesize()   [find huffman code sizes]
    Do 'do loop for (#non-zero-frequencies - 1) times
        V1 = -1                            'find highest v1 for      least value of freq(v1)>0
        V2 = -1                            'find highest v2 for next least value of freq(v2)>0
        swp = 2147483647 'Max Long variable
        swp2 = 2147483647
        For i = 0 To 256
            If freq(i) <> 0 Then
                If (freq(i) <= swp2) Then
                    If (freq(i) <= swp) Then
                        swp2 = swp
                        V2 = V1
                        swp = freq(i)
                        V1 = i
                    Else
                        swp2 = freq(i)
                        V2 = i
                    End If
                End If
            End If
        Next i
        If V2 = -1 Then
            freq(V1) = 0 'all elements in freq are now set to zero
            Exit Do      'done
        End If
        freq(V1) = freq(V1) + freq(V2)     'merge the two branches
        freq(V2) = 0
        codesize(V1) = codesize(V1) + 1    'Increment all codesizes in v1's branch
        While (others(V1) >= 0)
            V1 = others(V1)
            codesize(V1) = codesize(V1) + 1
        Wend
        others(V1) = V2                    'chain v2 onto v1's branch
        codesize(V2) = codesize(V2) + 1    'Increment all codesizes in v2's branch
        While (others(V2) >= 0)
            V2 = others(V2)
            codesize(V2) = codesize(V2) + 1
        Wend
    Loop

   'Count BITS  [find the number of codes of each size]
    n = 0
    For i = 0 To 256
        If codesize(i) <> 0 Then
            Bits(codesize(i)) = Bits(codesize(i)) + 1
            If n < codesize(i) Then n = codesize(i)    'Keep track of largest codesize
        End If
    Next i

   'Adjust BITS  [limit code lengths to 16 bits]
    i = n
    While i > 16
        While Bits(i) > 0
            For j = i - 2 To 1 Step -1        'Since symbols are paired for the longest Huffman
                If Bits(j) > 0 Then Exit For  'code, the symbols are removed from this length
            Next j                            'category two at a time.  The prefix for the pair
            Bits(i) = Bits(i) - 2             '(which is one bit shorter) is allocated to one
            Bits(i - 1) = Bits(i - 1) + 1     'of the pair;  then, (skipping the BITS entry for
            Bits(j + 1) = Bits(j + 1) + 2     'that prefix length) a code word from the next
            Bits(j) = Bits(j) - 1             'shortest non-zero BITS entry is converted into
        Wend                                  'a prefix for two code words one bit longer.
        i = i - 1
    Wend
    Bits(i) = Bits(i) - 1                  'Remove dummy symbol code from the code length count

   'Copy BITS and HUFFVAL to the HUFFMANTABLE  [HUFFVAL sorted by code length, then by value]
    With TheHuff
        For i = 1 To 16
            .Bits(i - 1) = Bits(i)
        Next i
        k = 0
        For i = 1 To n
            For j = 0 To 255
                If codesize(j) = i Then
                    .HUFFVAL(k) = j
                    k = k + 1
                End If
            Next j
        Next i
    End With

End Sub
Private Sub ExpandHuffman(TheHuff As HUFFMANTABLE, Optional MaxSymbol As Long = 255)
'Given a HUFFMANTABLE with valid BITS and HUFFVAL, generate tables for
'EHUFCO, EHUFSI, MAXCODE, and MINCODE so the table may be used for compression
'and/or decompression.  In JPEG, MaxSymbol is 255 for an AC Huffman Table.  For
'DC Tables, MaxSymbol is 11 for PP=8 bit precission, or 15 for PP=12 bit precission.
    Dim i          As Long 'Index for BITS
    Dim j          As Long 'Index for HUFFVAL
    Dim k          As Long 'Index for last HUFFVAL of length (i+1)
    Dim si         As Long 'Huffman code size  ( =2^i )
    Dim code       As Long 'Huffman code
    Dim symbol     As Long 'Huffman symbol


    With TheHuff

        For i = 0 To 255
           .EHUFSI(i) = 0      'Clear existing values so we can
           .EHUFCO(i) = -1     'check for duplicate huffman symbols
        Next i

        j = 0
        si = 1
        code = 0
        For i = 0 To 15
            k = j + .Bits(i)
            If k > 256 Then Err.Raise 1, , "Bad Huffman Table" 'more than 256 symbols
            If j = k Then 'no codes of length i+1
               .MINCODE(i) = j - code
               .MAXCODE(i) = -1
            Else
               .MINCODE(i) = j - code
                While j < k
                    symbol = .HUFFVAL(j)  'read symbol, make sure it's valid
                    If symbol > MaxSymbol Then Err.Raise 1, , "Bad Huffman Table"   'invalid symbol
                    If .EHUFCO(symbol) >= 0 Then Err.Raise 1, , "Bad Huffman Table" 'duplicate symbol
                   .EHUFSI(symbol) = si    'assign code for symbol
                   .EHUFCO(symbol) = code
                    code = code + 1
                    j = j + 1
                Wend
               .MAXCODE(i) = code - 1
            End If
            si = si * 2
            If code >= si Then Err.Raise 1, , "Bad Huffman Table" 'code does not fit into available bits
            code = code * 2
        Next i
        If j = 0 Then Err.Raise 1, , "Bad Huffman Table" 'No huffman symbols???
    End With

End Sub




'================================================================================
'                           E N T R O P Y   C O D I N G
'================================================================================
Private Sub WriteBitsBegin()
    m_Chr = 0
    m_Bit = 128
End Sub
Private Sub WriteBitsEnd()
    If m_Bit <> 128 Then WriteBits m_Bit, -1
End Sub
Private Sub WriteBits(ByVal si As Long, code As Long)
    While si > 0
        If (code And si) <> 0 Then m_Chr = (m_Chr Or m_Bit)
        If m_Bit = 1 Then            'We completed a byte ...
            m_Data(m_Ptr) = m_Chr    '    add it to the stream
            If m_Chr = 255 Then      'Pad a zero byte and advance pointer
                m_Data(m_Ptr + 1) = 0
                m_Ptr = m_Ptr + 2
            Else                     'just advance pointer
                m_Ptr = m_Ptr + 1
            End If
            m_Chr = 0                'clear byte buffer and reset bit index
            m_Bit = 128
        Else                         'increment to next bit position to write
            m_Bit = m_Bit \ 2
        End If
        si = si \ 2
    Wend
End Sub

Private Sub EncodeCoefficients(Data() As Integer, p As Long, Pred As Long, Td As Long, Ta As Long)
'Use Huffman tables to compress a block of 64 quantized DCT coefficients to the local
'm_Data() byte array.  The coefficients are input in the data() array starting at index p.
'Pred is the predictor for the DC coefficient.  Td and Ta are indexes to the local DC and AC
'Huffman Tables to use.
    Dim r     As Long
    Dim rs    As Long
    Dim si    As Long
    Dim code  As Long
    Dim p2    As Long

    p2 = p + 64

    code = Data(p) - Pred
    Pred = Data(p)
    p = p + 1

    si = 1
    rs = 0
    If code < 0 Then
        Do While si <= -code
            si = si * 2
            rs = rs + 1
        Loop
        code = code - 1
    Else
        Do While si <= code
            si = si * 2
            rs = rs + 1
        Loop
    End If
    si = si \ 2
    WriteBits HuffDC(Td).EHUFSI(rs), HuffDC(Td).EHUFCO(rs) 'append symbol for size category
    WriteBits si, code                                     'append diff

    With HuffAC(Ta)
        r = 0
        Do
            If Data(p) = 0 Then
                 r = r + 1
            Else
                While r > 15
                    WriteBits .EHUFSI(240), .EHUFCO(240) 'append RUN16 (a run of 16 zeros)
                    r = r - 16
                Wend
                code = Data(p)
                rs = r * 16
                si = 1
                If code < 0 Then
                    Do While si <= -code
                        si = si * 2
                        rs = rs + 1
                    Loop
                    code = code - 1
                Else
                    Do While si <= code
                        si = si * 2
                        rs = rs + 1
                    Loop
                End If
                si = si \ 2
                WriteBits .EHUFSI(rs), .EHUFCO(rs) 'append run length, size category
                WriteBits si, code                 'append AC value
                r = 0
            End If
            p = p + 1
        Loop While p < p2 'should be equal on exit
        If r <> 0 Then WriteBits .EHUFSI(0), .EHUFCO(0) 'append EOB (end of block)
    End With

End Sub




'========================================================================================
'                      C O L L E C T I N G   S T A T I S T I C S
'========================================================================================
'These procedures collect statistics of run-length and size categories of DCT coefficients
'so optimized Huffman tables can be generated to compress them.
Private Sub CollectStatisticsAC(Data() As Integer, freqac() As Long)
    Dim code As Long
    Dim n    As Long 'Number of coefficients in data()
    Dim p    As Long 'Index for current data() coefficient
    Dim p2   As Long
    Dim r    As Long 'Run length of zeros
    Dim rs   As Long 'Run-length/Size-category Symbol


    n = UBound(Data) + 1
    p = 0
    While p <> n
        p = p + 1     'Skip DC coefficient
        p2 = p + 63   '63 AC coefficients

        r = 0
        While p <> p2
            If Data(p) = 0 Then
                 r = r + 1
            Else
                While r > 15
                    freqac(240) = freqac(240) + 1  'RUN16 Symbol
                    r = r - 16
                Wend
                code = Data(p)
                If code < 0 Then 'rs = number of bits needed for code
                    rs = Int((Log(-code) * 1.442695040889)) + 1   '1/log(2)  (+ error correction)
                ElseIf code > 0 Then
                    rs = Int((Log(code) * 1.442695040889)) + 1    '1/log(2)  (+ error correction)
                Else
                    rs = 0
                End If

                rs = (r * 16) Or rs
                freqac(rs) = freqac(rs) + 1        'Run-length/Size-category Symbol
                r = 0
            End If
            p = p + 1
        Wend
        If r <> 0 Then freqac(0) = freqac(0) + 1   'EOB Symbol
    Wend

End Sub
Private Sub CollectStatisticsDCNonInterleaved(Data() As Integer, freqdc() As Long)
    Dim Diff  As Long     'DC Difference
    Dim Pred  As Long     'DC Predictor
    Dim n     As Long     'Number of coefficients in data()
    Dim p     As Long     'Index for current data() coefficient
    Dim s     As Long     'Size category for Diff


    n = UBound(Data) + 1
    p = 0
    Pred = 0
    While p <> n
        Diff = Data(p) - Pred
        Pred = Data(p)

        If Diff < 0 Then 's = number of bits needed for Diff
            s = Int((Log(-Diff) * 1.442695040889)) + 1   '1/log(2)  (+ error correction)
        ElseIf Diff > 0 Then
            s = Int((Log(Diff) * 1.442695040889)) + 1    '1/log(2) + (error correction)
        Else
            s = 0
        End If

        freqdc(s) = freqdc(s) + 1
        p = p + 64
    Wend

End Sub
Private Sub CollectStatisticsDCInterleaved(Data() As Integer, freqdc() As Long, Hi As Long, Vi As Long)
    Dim p()       As Long     'Index to .data in component f for scanline g
    Dim g         As Long      'Index counter  (sampling factor, vertical)
    Dim h         As Long      'Index counter  (sampling factor, horizontal)
    Dim i         As Long      'Index counter  (MCU horizontal)
    Dim j         As Long      'Index counter  (MCU vertical)
    Dim n         As Long      'Number of coefficients in data()
    Dim s         As Long      'Size category for Diff
    Dim Diff      As Long      'DC Difference
    Dim Pred      As Long      'DC Predictor
    Dim pLF       As Long      'Line Feed for p in data
    Dim MCUr      As Long      'Number of complete 8X8 blocks in rightmost MCU
    Dim MCUx      As Long      'Number of MCUs per scanline
    Dim MCUy      As Long      'Number of MCU scanlines


    n = UBound(Data) + 1
    ReDim p(Vi - 1)


    MCUx = (XX + 8 * HMax - 1) \ (8 * HMax)
    MCUy = (YY + 8 * VMax - 1) \ (8 * VMax)

    h = (-Int(-XX * Hi / HMax) + 7) \ 8  'Width of scanline in data (MCUs)

    For g = 0 To Vi - 1                  'Initialize .data pointers
        p(g) = 64 * h * g
    Next g
    pLF = 64 * h * (Vi - 1)              'Initialize .data pointer advancer

    MCUr = (h Mod Hi)                    'Number of complete 8X8 Blocks in rightmost MCU
    If MCUr = 0 Then MCUr = Hi

    For j = 1 To MCUy - 1

       'MCUs across a scanline
        For i = 1 To MCUx - 1
        For g = 1 To Vi
        For h = 1 To Hi

        Diff = Data(p(g - 1)) - Pred
        Pred = Data(p(g - 1))
        p(g - 1) = p(g - 1) + 64
        If Diff < 0 Then 's = number of bits needed for Diff
            s = Int((Log(-Diff) * 1.442695040889)) + 1   '1/log(2)  (+ error correction)
        ElseIf Diff > 0 Then
            s = Int((Log(Diff) * 1.442695040889)) + 1    '1/log(2) + (error correction)
        Else
            s = 0
        End If
        freqdc(s) = freqdc(s) + 1

        Next h
        Next g
        Next i

       'Rightmost MCU
        For g = 1 To Vi
        For h = 1 To Hi
        If h > MCUr Then  'Pad with dummy block
            s = 0
        Else
            Diff = Data(p(g - 1)) - Pred
            Pred = Data(p(g - 1))
            p(g - 1) = p(g - 1) + 64

            If Diff < 0 Then
                s = Int((Log(-Diff) * 1.442695040889)) + 1
            ElseIf Diff > 0 Then
                s = Int((Log(Diff) * 1.442695040889)) + 1
            Else
                s = 0
            End If
        End If
        freqdc(s) = freqdc(s) + 1
        Next h
        Next g

       'Advance data pointers
        For g = 0 To Vi - 1
            p(g) = p(g) + pLF
        Next g
    Next j


   'Bottommost MCU Scanline
    For i = 1 To MCUx
    For g = 1 To Vi
    For h = 1 To Hi
    If p(g - 1) >= n Or (i = MCUx And h > MCUr) Then 'Pad with dummy block
        s = 0
    Else
        Diff = Data(p(g - 1)) - Pred
        Pred = Data(p(g - 1))
        p(g - 1) = p(g - 1) + 64

        If Diff < 0 Then
            s = Int((Log(-Diff) * 1.442695040889)) + 1
        ElseIf Diff > 0 Then
            s = Int((Log(Diff) * 1.442695040889)) + 1
        Else
            s = 0
        End If
    End If
    freqdc(s) = freqdc(s) + 1
    Next h
    Next g
    Next i

End Sub




'========================================================================================
'                                Q U A N T I Z A T I O N
'========================================================================================
Private Sub ExpandDQT(Tqi As Long)
    Dim i          As Long
    Dim j          As Long
    Dim k          As Byte
    Dim maxvalue   As Long

    With QTable(Tqi)
        If PP = 12 Then
            maxvalue = 65535
        Else
            maxvalue = 255
        End If

        For i = 0 To 7
            For j = 0 To 7
                k = ZigZag(i, j)
                If .Qk(k) < 1 Or .Qk(k) > maxvalue Then Err.Raise 1, , "Bad Quantization Table"
               .FScale(k) = FDCTScale(i) * FDCTScale(j) / CDbl(.Qk(k))
            Next j
        Next i
    End With

End Sub
Private Sub Quantize(Data() As Integer, p As Long, FScale() As Single)
    Dim i As Long  'Take 8X8 block of unscaled DCT coefficients [m_Block(0-7, 0-7)],
    Dim j As Long  'Scale, Quantize, and store the results in data() array of
    Dim t As Long  'COMPONENT in Zig Zag order at index p


    For j = 0 To 7
        For i = 0 To 7
            t = ZigZag(i, j)
            Data(p + t) = m_Block(i, j) * FScale(t)
        Next i
    Next j
    p = p + 64

End Sub
Public Property Let Quality(vData As Long)
'The JPEG compression standard does not have a formal definition for image Quality.
'This implementation defines Quality as an integer value between 1 and 100, and
'generates quantization tables based on the value given.
'
'Quality < 50  -  Poor image quality with high compression
'Quality = 75  -  Good quality pictures for displaying on a monitor or web page ... typical for general use
'Quality = 92  -  High quality with non-optimal compression ... Appropriate for printing ... [typical digital camera "max quality" setting]
'Quality > 95  -  Wasteful ... very poor compression with little image quality improvement.  Use 24-bit BMP TrueColor if you need quality this high.

    Dim i           As Long
    Dim qvalue      As Long
    Dim maxvalue    As Long
    Dim scalefactor As Long

    maxvalue = 255 '32767 if 16 bit quantum is allowed

    If vData > 0 And vData <= 100 Then
        m_Quality = vData

        If (m_Quality < 50) Then
            If m_Quality <= 0 Then
                scalefactor = 5000
            Else
                scalefactor = 5000 / m_Quality
            End If
        Else
            If m_Quality > 100 Then
                scalefactor = 0
            Else
                scalefactor = 200 - m_Quality * 2
            End If
        End If

        With QTable(0)
            For i = 0 To 63
                qvalue = (QLumin(i) * scalefactor + 50) / 100
                If qvalue <= 0 Then
                    qvalue = 1
                ElseIf qvalue > maxvalue Then
                    qvalue = maxvalue
                End If
                .Qk(i) = qvalue
            Next i
        End With
        With QTable(1)
            For i = 0 To 63
                qvalue = (QChrom(i) * scalefactor + 50) / 100
                If qvalue <= 0 Then
                    qvalue = 1
                ElseIf qvalue > maxvalue Then
                    qvalue = maxvalue
                End If
                .Qk(i) = qvalue
            Next i
        End With

        ExpandDQT 0
        ExpandDQT 1
    End If

End Property
Public Property Get Quality() As Long
    Quality = m_Quality
End Property





'================================================================================
'                           I M A G E   S A M P L I N G
'================================================================================
Public Sub SetSamplingFrequencies(H1 As Long, V1 As Long, H2 As Long, V2 As Long, H3 As Long, V3 As Long)

'This class always samples and compresses pictures in YCbCr colorspace.  The first component, Y,
'represents the Luminance of the pixels.  This is "how bright" a pixel is.  The Cb and Cr
'components are Chrominance, which is a measure of how far from neutral-white (toward a color)
'a pixel is.  The human visual sensory system can discriminate Luminance differences about
'twice as well as it can discriminate Chrominance differences.
'
'Virtually all JPEG files are in YCbCr colorspace.  Other JPEG compliant colorspaces exist, but
'they are used in specialty equipment.  For example, people in the astronomy or medical fields
'choose colorspaces that best record the information they are interested in, and don't care about
'how pretty the picture looks to a person when displayed on a computer monitor.
'[Apple/Machintosh sometimes uses a four component colorspace, but that colorspace is rare and
'not widely supported]
'
'Sampling frequencies define how often each component is sampled.  Higher frequencies store more
'information, while lower frequencies store less.  Typically, sampling frequencies are set at
'2,2, 1,1, 1,1.  This corresponds to the human visual sensory system.  The first component,
'Luminance, is sampled twice as much because our eyes notice differences in Luminance quite easily.
'The two Chrominance components are sampled half as much as because our eyes can't distinguish
'the difference in color changes as well.  One Luminance value is sampled for every pixel, and
'one Chrominance value is sampled for each 2X2 block of pixels.
'
'Digital cameras typically record at sampling frequencies of 1,1, 1,1, 1,1.  This samples every
'pixel for all three components.  The quality of the picture is a little better when viewed by
'a person, but the compression benefits drop significantly.  If the picture to be compressed
'is from a Scanner or Digital camera, and you plan on printing it in the future, and storage
'space is not a problem, then sampling at these frequencies makes sense.  Otherwise, if you only
'plan on using the picture to display on a monitor or a web page, [2,2, 1,1, 1,1] makes the
'most sense.
'
'The JPEG standard specifies that sampling frequencies may range from 1-4 for each component
'in both directions.  However, if any component has a sampling frequency of '3', and another
'component has a coresponding sampling frequency of '2' or '4', the downsampling process
'will map fractional pixels to sample values.  This is leagal in the JPEG standard, and this
'class will compress fractional pixel samplings, but this is not widely supported.  It is
'highly recommended to AVOID SAMPLING FACTORS OF 3 for maximum compatability with JPEG decoders.
'
'Some JPEG encoders avoid the fractional pixel problem by only allowing the end user to pick
'a "sub-sampling" value.  In such "Sub Sampling" schemes, all Chrominance frequencies are set
'to one, and the (one or two) sub-sampling value(s) specify Luminance frequencies.
'
'There should *never* be an error raised if you are using this class correctly.  It should
'not be possible for the end user to specify illegal sampling frequency values!
'[For tinkerers - If you delete the error raising code and specify illegal sampling
'frequencies, this class will procede to create a non-JPEG compliant file with the values
'specified]

    Dim i As Long

    If H1 < 1 Or H1 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
    If V1 < 1 Or V1 > 4 Then Err.Raise 1, , "Invalid Sampling Value"

    If (H2 Or H3 Or V2 Or V3) = 0 Then  'if H2,H3,V2,V3 are all zero ...
        Nf = 1         'Luminance only.
        ReDim Comp(0)
        Comp(0).Hi = 1 'Set up for sampling Greyscale
        Comp(0).Vi = 1 '(Black and White picture)
    Else
        If H2 < 1 Or H2 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
        If H3 < 1 Or H3 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
        If V2 < 1 Or V2 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
        If V3 < 1 Or V3 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
        Nf = 3         'YCbCr
        ReDim Comp(2)
        Comp(0).Hi = H1
        Comp(0).Vi = V1
        Comp(0).Tqi = 0
        Comp(1).Hi = H2
        Comp(1).Vi = V2
        Comp(1).Tqi = 1
        Comp(2).Hi = H3
        Comp(2).Vi = V3
        Comp(2).Tqi = 1
    End If

    HMax = -1
    VMax = -1
    For i = 0 To Nf - 1 'determine max h, v sampling factors
        If HMax < Comp(i).Hi Then HMax = Comp(i).Hi
        If VMax < Comp(i).Vi Then VMax = Comp(i).Vi
    Next i

End Sub

Public Function SampleHDC(ByVal lHDC As Long, lWidth As Long, lHeight As Long, Optional lSrcLeft As Long, Optional lSrcTop As Long) As Long
'Given a valid hDC and dimensions, generate component samplings of an Image.
'A DIBSection is created to hold Sample(s) of the Image, from which the Image is
'decomposed into YCbCr components.
'Returns: 0 = Success
'         1 = API error while generating a DIBSection
    Dim hDIB       As Long    'Handle to the DIBSection
    Dim hBmpOld    As Long    'Handle to the old bitmap in the DC, for clear up
    Dim hDC        As Long    'Handle to the Device context holding the DIBSection
    Dim lPtr       As Long    'Address of memory pointing to the DIBSection's bits
    Dim BI         As BITMAPINFO 'Type containing the Bitmap information
    Dim SA         As SAFEARRAY2D
    Dim Pixel()    As Byte   'Byte array containing pixel data
    Dim f          As Long   'Index counter for components
    Dim qp         As Long   'Index for quantized FDCT value (in component data)
    Dim rm         As Single 'Scale factor for red   pixel when converting RGB->YCbCr
    Dim gm         As Single 'Scale factor for green pixel when converting RGB->YCbCr
    Dim bm         As Single 'Scale factor for blue  pixel when converting RGB->YCbCr
    Dim s          As Single 'Level shift value             for converting RGB->YCbCr
    Dim xi         As Long   'Sample width
    Dim yi         As Long   'Sample height
    Dim xi2        As Long   'Sample width  (for previous component)
    Dim yi2        As Long   'Sample height (for previous component)
    Dim xi8        As Long   'Sample width  (padded to 8 pixel barrier)
    Dim yi8        As Long   'Sample height (padded to 8 pixel barrier)
    Dim i0         As Long   'Left index of an 8X8 block of pixels
    Dim j0         As Long   'Top  index of an 8X8 block of pixels
    Dim i          As Long   'Pixel Index (Horizontal)
    Dim j          As Long   'Pixel Index (Vertical)
    Dim p          As Long   'DCT Index (horizontal)
    Dim q          As Long   'DCT Index (vertical)


    PP = 8
    YY = lHeight
    XX = lWidth

   'Create a DIBSection to store Sampling(s) of the Image
    hDC = CreateCompatibleDC(0)
    If hDC = 0 Then
        SampleHDC = 1 'CreateCompatibleDC() API Failure
    Else
        With BI.bmiHeader
            .biSize = Len(BI.bmiHeader)
            .biWidth = (lWidth + 7) And &HFFFFFFF8   '8 byte barrier for 8X8 data units
            .biHeight = (lHeight + 7) And &HFFFFFFF8
            .biPlanes = 1
            .biBitCount = 24
            .biCompression = BI_RGB
            .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight '4 byte barrier
        End With
        hDIB = CreateDIBSection2(hDC, BI, DIB_RGB_COLORS, lPtr, 0, 0)
        If hDIB = 0 Then
            SampleHDC = 1 'CreateDIBSection2() API Failure
        Else
            With SA                        'This code copies the pointer of the 2-D bitmap
                .cbElements = 1            'pixel data to the pointer of the Pixel() array.
                .cDims = 2                 'This allows you to read/modify the pixel data
                .Bounds(0).lLbound = 0     'as if it were stored in the Pixel() array.
                .Bounds(0).cElements = BI.bmiHeader.biHeight
                .Bounds(1).lLbound = 0
                .Bounds(1).cElements = ((BI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC)
                .pvData = lPtr             'Note that this is extreamly efficient, since it copies
            End With                       'a pointer to the data, and not the data itself.
            hBmpOld = SelectObject(hDC, hDIB) 'Select DIBSection into DC
            If SetStretchBltMode(hDC, HALFTONE) = 0 Then SetStretchBltMode hDC, COLORONCOLOR


    For f = 0 To Nf - 1
        Select Case f 'Select scaling factors for RGB->YCbCr conversion for this component
        Case 0 'Luminance
            rm = 0.299
            gm = 0.587
            bm = 0.114
            s = -128
        Case 1 'Chrominance [Blue-Yellow]
            rm = -0.16874
            gm = -0.33126
            bm = 0.5
            s = 0
        Case 2 'Chrominance [Red-Green]
            rm = 0.5
            gm = -0.41869
            bm = -0.08131
            s = 0
        End Select

        With Comp(f)
           .Ci = f + 1 'Assign an ID to this component

            xi = -Int(-XX * .Hi / HMax)        'determine Sample dimensions
            yi = -Int(-YY * .Vi / VMax)
            xi8 = ((xi + 7) And &HFFFFFFF8)    'Sample dimensions with 8X8 barrier
            yi8 = ((yi + 7) And &HFFFFFFF8)
            ReDim .Data(xi8 * yi8 - 1)

            If xi8 <> xi2 Or yi8 <> yi2 Then  'We need to Sample the Image
                If xi = XX And yi = YY Then 'Just copy the image to our DIBSection
                    BitBlt hDC, 0, BI.bmiHeader.biHeight - yi8, xi, yi, lHDC, lSrcLeft, lSrcTop, vbSrcCopy
                Else                        'Resample/Resize the Image
                    StretchBlt hDC, 0, BI.bmiHeader.biHeight - yi8, xi, yi, lHDC, lSrcLeft, lSrcTop, lWidth, lHeight, vbSrcCopy
                End If
                For i = xi To xi8 - 1  'Pad right of Sample to 8 block barrier
                    BitBlt hDC, i, BI.bmiHeader.biHeight - yi8, 1, yi, hDC, i - 1, BI.bmiHeader.biHeight - yi8, vbSrcCopy
                Next i
                For j = BI.bmiHeader.biHeight - (yi8 - yi) To BI.bmiHeader.biHeight - 1 'Pad bottom of Sample to 8 block barrier
                    BitBlt hDC, 0, j, xi8, 1, hDC, 0, j - 1, vbSrcCopy
                Next j
            End If
            xi2 = xi8
            yi2 = yi8
            qp = 0 'Reset output Quantized FDCT Coefficient Index

           'Read 8X8 blocks of pixels, convert from RGB->YCbCr colorspace, FDCT and Quantize
           'the data, store the results in .data of this component
            CopyMemory ByVal VarPtrArray(Pixel), VarPtr(SA), 4& 'Get Pixel array descriptor
            j = yi8 - 1
            While j > 0               'Scan from top to bottom (j = -1 after loop)
                i = 0
                j0 = j
                While i < 3 * xi8     'Scan from left to right (i = 3*xi8 after loop)
                    j = j0
                    i0 = i
                    For p = 0 To 7    'Get 8X8 block of level shifted YCbCr values
                        i = i0
                        For q = 0 To 7
                            m_Block(q, p) = rm * Pixel(i + 2, j) + _
                                            gm * Pixel(i + 1, j) + _
                                            bm * Pixel(i, j) + s
                            i = i + 3
                        Next q
                        j = j - 1
                    Next p
                    FDCT                               'Calculate the FDCT
                    Quantize .Data, qp, QTable(.Tqi).FScale   'Quantize, and store in DCT buffer
                Wend
            Wend
            CopyMemory ByVal VarPtrArray(Pixel), 0&, 4 'Clear the Pixel array descriptor
        End With
    Next f


            SelectObject hDC, hBmpOld 'Select CompatibleDC  (unselect DIBSection)
            DeleteObject hDIB         'Delete DIBSection
        End If
        DeleteObject hDC              'Delete CompatibleDC
    End If

End Function



Public Property Let Comment(Value As String)
   'Assigning a value to this property will add the text Comment to the JPEG file.
    If Len(Value) > 65535 Then Err.Raise 1, , "Illegal Comment Length"
    m_Comment = Value
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property




'================================================================================
'                         E M I T I N G   M A R K E R S
'================================================================================
Private Sub InsertJFIF()
    If m_Ptr + 17 > UBound(m_Data) Then Err.Raise 9 'Copymemory will write past bounds of m_Data()

    CopyMemory m_Data(m_Ptr + 0), &H1000E0FF, 4&    'APP0 Marker, Length(APP0)=16
    CopyMemory m_Data(m_Ptr + 4), &H4649464A, 4&    '"JFIF"
    CopyMemory m_Data(m_Ptr + 8), &H10100, 4&       '"/0", Version Major=1, Version Minor=1
                                                    'Units=0  [0=pixel, 1=dpi, 2=dots/cm]
    CopyMemory m_Data(m_Ptr + 12), &H1000100, 4&    'Horizontal pixel density = 1 (dot per pixel)
                                                    'Vertical   pixel density = 1 (dot per pixel)
    CopyMemory m_Data(m_Ptr + 16), &H0&, 2&         'Thumbnail horizontal pixel count = 0
    m_Ptr = m_Ptr + 18                              'Thumbnail vertical   pixel count = 0

End Sub

Private Sub InsertSOF(SOFMarker As Long)
    Dim i   As Long 'Insert a Start Of Frame marker segment
    Dim Lx  As Long 'PP, YY, XX, Nf, and Ci,Hi,Vi,Tqi, must already be set

    Lx = 8 + 3 * Nf
    m_Data(m_Ptr) = 255                    'SOF
    m_Data(m_Ptr + 1) = SOFMarker And 255
    m_Data(m_Ptr + 2) = Lx \ 256           'Frame Header Length
    m_Data(m_Ptr + 3) = Lx And 255
    m_Data(m_Ptr + 4) = PP                 'Sample precision [8, 12]
    m_Data(m_Ptr + 5) = YY \ 256           'Number of Lines
    m_Data(m_Ptr + 6) = YY And 255
    m_Data(m_Ptr + 7) = XX \ 256           'Number of samples per line
    m_Data(m_Ptr + 8) = XX And 255
    m_Data(m_Ptr + 9) = Nf                 'Number of image components in frame
    m_Ptr = m_Ptr + 10
    For i = 0 To Nf - 1                      'For each component ...
        With Comp(i)
            m_Data(m_Ptr) = .Ci                  'Component identifier
            m_Data(m_Ptr + 1) = .Hi * 16 Or .Vi  'Horizontal/Vertical sampling factors
            m_Data(m_Ptr + 2) = .Tqi             'Quantization table selector
        End With
        m_Ptr = m_Ptr + 3
    Next i
End Sub

Private Sub InsertCOM(TheComment As String)
    Dim i As Long
    Dim Lx As Long

    Lx = Len(TheComment) + 2
    If Lx > 2 Then
        m_Data(m_Ptr) = 255               'COM marker
        m_Data(m_Ptr + 1) = COM
        m_Data(m_Ptr + 2) = Lx \ 256      'COM marker segment length
        m_Data(m_Ptr + 3) = Lx And 255
        m_Ptr = m_Ptr + 4
        For i = 1 To Len(TheComment)      'Comment text
            m_Data(m_Ptr) = Asc(Mid$(TheComment, i, 1))
            m_Ptr = m_Ptr + 1
        Next i
    End If
End Sub
Private Sub InsertDQT(ByVal MarkerPos As Long, Tqi As Long)
    Dim i As Long 'Call with MarkerPos = m_Ptr to insert a single table with its own DQT marker
                  'Call multiple times with the same MarkerPos to include
                  'multiple tables under the same DQT marker

    If m_Ptr < MarkerPos + 4 Then 'Insert Marker
        m_Ptr = MarkerPos + 4
        m_Data(m_Ptr - 4) = 255
        m_Data(m_Ptr - 3) = DQT
    End If
    With QTable(Tqi)
        For i = 0 To 63
            If .Qk(i) > 255 Then Exit For
        Next i
        If i = 64 Then              '8 bit precision
            m_Data(m_Ptr) = Tqi
            m_Ptr = m_Ptr + 1
            For i = 0 To 63
                m_Data(m_Ptr) = .Qk(i)
                m_Ptr = m_Ptr + 1
            Next i
        Else                        '16 bit precision
            If PP <> 12 Then Err.Raise 1, , "Illegal precission in Quantization Table"
            m_Data(m_Ptr) = Tqi Or 16
            m_Ptr = m_Ptr + 1
            For i = 0 To 63
                m_Data(m_Ptr) = .Qk(i) \ 256
                m_Data(m_Ptr + 1) = .Qk(i) And 255
                m_Ptr = m_Ptr + 2
            Next i
        End If
    End With

    m_Data(MarkerPos + 2) = (m_Ptr - MarkerPos - 2) \ 256& 'Insert Marker segment length
    m_Data(MarkerPos + 3) = (m_Ptr - MarkerPos - 2) And 255&
End Sub
Private Sub InsertDHT(ByVal MarkerPos As Long, HIndex As Long, IsAC As Boolean)
    Dim i As Long 'Call with MarkerPos = m_Ptr to insert a single table with its own DHT marker
    Dim j As Long 'Call multiple times with the same MarkerPos to include
                  'multiple tables under the same DHT marker

    If m_Ptr < MarkerPos + 4 Then 'Insert Marker
        m_Ptr = MarkerPos + 4
        m_Data(m_Ptr - 4) = 255
        m_Data(m_Ptr - 3) = DHT
    End If
    If IsAC Then
        With HuffAC(HIndex)
            m_Data(m_Ptr) = HIndex Or 16
            m_Ptr = m_Ptr + 1
            j = 0
            For i = 0 To 15
                m_Data(m_Ptr) = .Bits(i)
                m_Ptr = m_Ptr + 1
                j = j + .Bits(i)
            Next i
            For i = 0 To j - 1
                m_Data(m_Ptr) = .HUFFVAL(i)
                m_Ptr = m_Ptr + 1
            Next i
        End With
    Else
        With HuffDC(HIndex)
            m_Data(m_Ptr) = HIndex
            m_Ptr = m_Ptr + 1
            j = 0
            For i = 0 To 15
                m_Data(m_Ptr) = .Bits(i)
                m_Ptr = m_Ptr + 1
                j = j + .Bits(i)
            Next i
            For i = 0 To j - 1
                m_Data(m_Ptr) = .HUFFVAL(i)
                m_Ptr = m_Ptr + 1
            Next i
        End With
    End If

    m_Data(MarkerPos + 2) = (m_Ptr - MarkerPos - 2) \ 256& 'Insert Marker segment length
    m_Data(MarkerPos + 3) = (m_Ptr - MarkerPos - 2) And 255&
End Sub
Private Sub InsertMarker(TheMarker As Long)
    m_Data(m_Ptr) = 255
    m_Data(m_Ptr + 1) = TheMarker
    m_Ptr = m_Ptr + 2
End Sub




'================================================================================
'                           E M I T I N G   S C A N S
'================================================================================
Private Sub InsertSOSNonInterleaved(CompIndex As Long, Td As Long, Ta As Long)
'Insert an SOS marker and scan data for a non-interleaved Sequential scan.
    Dim p         As Long     'Index to .data in component
    Dim n         As Long
    Dim Pred      As Long     'Predictor for DC coefficient


   'Insert SOS Marker Segment
    m_Data(m_Ptr) = 255                          'SOS Marker
    m_Data(m_Ptr + 1) = SOS
    m_Data(m_Ptr + 2) = 8 \ 256                  'Marker Segment Length
    m_Data(m_Ptr + 3) = 8 And 255
    m_Data(m_Ptr + 4) = 1                        'Ns     - Number of components in Scan [1-4]
    m_Ptr = m_Ptr + 5
    m_Data(m_Ptr) = Comp(CompIndex).Ci           'Csj    - Component ID
    m_Data(m_Ptr + 1) = Td * 16 Or Ta            'Td, Ta - DC, AC entropy coder selector
    m_Ptr = m_Ptr + 2
    m_Data(m_Ptr) = 0                            'Ss     - Start of spectral selection
    m_Data(m_Ptr + 1) = 63                       'Se     - End of spectral selection
    m_Data(m_Ptr + 2) = 0                        'Ah, Al - Successive approximation bit high/low
    m_Ptr = m_Ptr + 3

   'Insert non-interleaved sequential entropy coded data
    With Comp(CompIndex)

        p = 0
        n = UBound(.Data) + 1
        Pred = 0

        WriteBitsBegin
        While p <> n
            EncodeCoefficients .Data, p, Pred, Td, Ta
        Wend
        WriteBitsEnd

    End With

End Sub

Private Sub InsertSOSInterleaved(CompIndex() As Long, Td() As Long, Ta() As Long, FirstIndex As Long, SecondIndex As Long)
'Insert an SOS marker and scan data for an interleaved Sequential scan.

    Dim f         As Long      'Index counter  (component)
    Dim g         As Long      'Index counter  (sampling factor, vertical)
    Dim h         As Long      'Index counter  (sampling factor, horizontal)
    Dim i         As Long      'Index counter  (MCU horizontal)
    Dim j         As Long      'Index counter  (MCU vertical)
    Dim Lx        As Long      'Marker Segment Length
    Dim Ns        As Long      'Number of components in Scan [1-4]
    Dim MCUx      As Long      'Number of MCUs per scanline
    Dim MCUy      As Long      'Number of MCU scanlines

    Dim p()        As Long     'Index to .data in component f for scanline g
    Dim pLF()      As Long     'Line Feed for p in .data for component f
    Dim Pred()     As Long     'Predictor for DC coefficient in component f
    Dim MCUr()     As Long     'Number of complete 8X8 blocks in rightmost MCU for component f
    Dim Pad64(63)  As Integer  '8X8 padding block for completing MCUs


    Ns = SecondIndex - FirstIndex + 1
    Lx = 6 + 2 * Ns

   'Insert SOS Marker Segment
    m_Data(m_Ptr) = 255                          'SOS Marker
    m_Data(m_Ptr + 1) = SOS
    m_Data(m_Ptr + 2) = Lx \ 256                 'Marker Segment Length
    m_Data(m_Ptr + 3) = Lx And 255
    m_Data(m_Ptr + 4) = Ns                       'Ns     - Number of components in Scan [1-4]
    m_Ptr = m_Ptr + 5
    For i = FirstIndex To SecondIndex
        m_Data(m_Ptr) = Comp(CompIndex(i)).Ci   'Csj
        m_Data(m_Ptr + 1) = Td(i) * 16 Or Ta(i) 'Td, Ta
        m_Ptr = m_Ptr + 2
    Next i
    m_Data(m_Ptr) = 0                            'Ss     - Start of spectral selection
    m_Data(m_Ptr + 1) = 63                       'Se     - End of spectral selection
    m_Data(m_Ptr + 2) = 0                        'Ah, Al - Successive approximation bit high/low
    m_Ptr = m_Ptr + 3


   'Insert interleaved sequential entropy coded data
    ReDim p(FirstIndex To SecondIndex, VMax - 1)
    ReDim Pred(FirstIndex To SecondIndex)
    ReDim pLF(FirstIndex To SecondIndex)
    ReDim MCUr(FirstIndex To SecondIndex)

    MCUx = (XX + 8 * HMax - 1) \ (8 * HMax)
    MCUy = (YY + 8 * VMax - 1) \ (8 * VMax)

    For f = FirstIndex To SecondIndex
        With Comp(CompIndex(f))
            h = (-Int(-XX * .Hi / HMax) + 7) \ 8  'Width of scanline in .data (MCUs)

            For g = 0 To .Vi - 1                  'Initialize .data pointers
                p(f, g) = 64 * h * g
            Next g
            pLF(f) = 64 * h * (.Vi - 1)           'Initialize .data pointer advancer

            MCUr(f) = (h Mod .Hi)                 'Number of complete 8X8 Blocks in rightmost MCU
            If MCUr(f) = 0 Then MCUr(f) = .Hi
        End With
    Next f

    WriteBitsBegin
    For j = 1 To MCUy - 1

       'Encode MCUs across a scanline
        For i = 1 To MCUx - 1
        For f = FirstIndex To SecondIndex '0 To Ns - 1
        With Comp(CompIndex(f))
        For g = 1 To .Vi
        For h = 1 To .Hi
        EncodeCoefficients .Data, p(f, g - 1), Pred(f), Td(f), Ta(f)
        Next h
        Next g
        End With
        Next f
        Next i

       'Encode Rightmost MCU
        For f = FirstIndex To SecondIndex '0 To Ns - 1
        With Comp(CompIndex(f))
        For g = 1 To .Vi
        For h = 1 To .Hi
        If h > MCUr(f) Then 'Pad with dummy block
            Pad64(0) = Pred(f)
            EncodeCoefficients Pad64, 0, Pred(f), Td(f), Ta(f)
        Else
            EncodeCoefficients .Data, p(f, g - 1), Pred(f), Td(f), Ta(f)
        End If
        Next h
        Next g
        End With
        Next f

       'Advance .data pointers
        For f = FirstIndex To SecondIndex
        For g = 0 To Comp(CompIndex(f)).Vi - 1
        p(f, g) = p(f, g) + pLF(f)
        Next g
        Next f
     Next j

   'Encode Bottommost MCU Scanline
    For i = 1 To MCUx
    For f = FirstIndex To SecondIndex
    With Comp(CompIndex(f))
    For g = 1 To .Vi
    For h = 1 To .Hi
    If p(f, g - 1) > UBound(.Data) Or (i = MCUx And h > MCUr(f)) Then 'Pad with dummy block
        Pad64(0) = Pred(f)
        EncodeCoefficients Pad64, 0, Pred(f), Td(f), Ta(f)
    Else
        EncodeCoefficients .Data, p(f, g - 1), Pred(f), Td(f), Ta(f)
    End If
    Next h
    Next g
    End With
    Next f
    Next i

    WriteBitsEnd

End Sub

Private Sub InsertSequentialScans(CompIndex() As Long, Td() As Long, Ta() As Long, FirstIndex As Long, SecondIndex As Long)
'Insert scan components CompIndex(FirstIndex) to CompIndex(SecondIndex) sequentially in compliance
'with JPEG rules.  Components are interleaved whenever possible to emit as few scans as possible.

    Dim f            As Long       'First Index
    Dim g            As Long       'Second Index
    Dim Nb           As Long       'Number of 8X8 blocks in MCU
    Const MaxNb      As Long = 10  'Max 8X8 blocks in MCU  (10 for JPEG compliance)
    Dim flag         As Boolean    'True when ready to insert scan(s)

    f = FirstIndex
    g = FirstIndex
    Nb = 0
    flag = False
    While f <= SecondIndex

        Nb = Nb + Comp(CompIndex(g)).Hi * Comp(CompIndex(g)).Vi
        g = g + 1

        If Nb > MaxNb Then
            flag = True
            If f <> g - 1 Then g = g - 1
        Else
            If (g - f) = 3 Or g > SecondIndex Then flag = True
        End If

        If flag Then
            If f = g - 1 Then
                InsertSOSNonInterleaved CompIndex(f), Td(f), Ta(f)
            Else
                InsertSOSInterleaved CompIndex, Td, Ta, f, g - 1
            End If
            Nb = 0
            f = g
            flag = False
        End If
    Wend

End Sub





'========================================================================================
'                               W R I T I N G   F I L E
'========================================================================================
Private Function OptimizeHuffmanTables(CompIndex() As Long, Td() As Long, Ta() As Long, FirstIndex As Long, SecondIndex As Long) As Long
'Optimize Huffman tables for the component indexes given.
'Returns an estimate of the number of bytes needed for entropy coded data.
'Estimate assumes a single scan, and entropy coded FF bytes are not followed by a zero stuff byte.

    Dim f            As Long      'First Index
    Dim g            As Long      'Second Index
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long      'Total bytes required for entropy coded data
    Dim k1           As Long
    Dim k2           As Long
    Dim Nb           As Long      'Number of 8X8 blocks in MCU
    Const MaxNb      As Long = 10 'Max 8X8 blocks in MCU  (10 for JPEG compliance)
    Dim freq(256)    As Long      'frequency count for optimizing Huffman tables
    Dim freq2()      As Long      'copy of freq, used for calcultating entropy coded data size
    Dim IsInter()    As Boolean   'True if component i will be interleaved
    Dim TdUsed()     As Boolean   'True if HuffDC(i) is used
    Dim TaUsed()     As Boolean   'True if HuffAC(i) is used
    Dim flag         As Boolean   'True when ready to include scan(s)

    ReDim IsInter(FirstIndex To SecondIndex)
    ReDim TaUsed(3)
    ReDim TdUsed(3)


   'Determine which components will be interleaved by InsertSequentialScans(), which tables are used
    f = FirstIndex
    g = FirstIndex
    Nb = 0
    flag = False
    While f <= SecondIndex

        Nb = Nb + Comp(CompIndex(g)).Hi * Comp(CompIndex(g)).Vi
        g = g + 1

        If Nb > MaxNb Then
            flag = True
            If f <> g - 1 Then g = g - 1
        Else
            If (g - f) = 3 Or g > SecondIndex Then flag = True
        End If

        If flag Then
            If f = g - 1 Then
                    TdUsed(Td(f)) = True
                    TaUsed(Ta(f)) = True
                    IsInter(f) = False
            Else
                For i = f To g - 1
                    TdUsed(Td(i)) = True
                    TaUsed(Ta(i)) = True
                    IsInter(i) = True
                Next i
            End If
            Nb = 0
            f = g
            flag = False
        End If
    Wend


   'Optimize huffman tables for the scan sequence
    For i = 0 To 3
        If TdUsed(i) Then
            For f = FirstIndex To SecondIndex
                With Comp(CompIndex(f))
                    If Td(f) = i Then
                        If IsInter(f) Then
                            CollectStatisticsDCInterleaved .Data, freq, .Hi, .Vi
                        Else
                            CollectStatisticsDCNonInterleaved .Data, freq
                        End If
                    End If
                End With
            Next f

           'Optimize and create this DC table
            freq2 = freq
            OptimizeHuffman HuffDC(i), freq
            ExpandHuffman HuffDC(i), IIf(PP = 12, 15, 11)

           'Calculate compressed data size and add to total k
            For j = 0 To 15
                If freq2(j) <> 0 Then
                    k1 = j + Int(Log(HuffDC(i).EHUFSI(j)) * 1.442695040889) + 1 'Length of size category, in bits
                    k2 = k2 + freq2(j) * k1                                     'Sum all occurances of this coefficient, in bits
                    k = k + k2 \ 8                                              'add to byte count
                    k2 = k2 Mod 8                                               'preserve remaining bits
                End If
            Next j

        End If
        If TaUsed(i) Then
            For f = FirstIndex To SecondIndex
                If Td(f) = i Then CollectStatisticsAC Comp(CompIndex(f)).Data, freq
            Next f

           'Optimize and create this AC table
            freq2 = freq
            OptimizeHuffman HuffAC(i), freq
            ExpandHuffman HuffAC(i), 255

           'Calculate compressed data size and add to total k
            For j = 0 To 255
                If freq2(j) <> 0 Then
                    k1 = (j And 15) + Int(Log(HuffAC(i).EHUFSI(j)) * 1.442695040889) + 1 'Length of size category, in bits
                    k2 = k2 + freq2(j) * k1                                              'Sum all occurances of this coefficient, in bits
                    k = k + k2 \ 8                                                       'add to byte count
                    k2 = k2 Mod 8                                                        'preserve remaining bits
                End If
            Next j

        End If
    Next i

    If (k2 Mod 8) <> 0 Then k = k + 1
    OptimizeHuffmanTables = k

End Function



Public Function SaveFile(FileName As String) As Long
    Dim CompIndex()  As Long 'Indexes of Components to be included
    Dim Td()         As Long 'DC Huffman Table Selectors
    Dim Ta()         As Long 'AC Huffman Table Selectors
    Dim fileNum      As Integer
    Dim i            As Long
    
    Message "Preparing headers..."


    If Len(FileName) = 0 Then
        SaveFile = 1         'FileName not given
    Else
        If (Len(Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0) Then
            SaveFile = 2     'File already exists
        Else


    ReDim CompIndex(Nf - 1)
    ReDim Td(Nf - 1)
    ReDim Ta(Nf - 1)

    For i = 0 To Nf - 1
        CompIndex(i) = i
        Td(i) = IIf(i = 0, 0, 1)
        Ta(i) = IIf(i = 0, 0, 1)
    Next i

    Message "Preparing compression tables..."

    i = OptimizeHuffmanTables(CompIndex, Td, Ta, 0, Nf - 1)

   'Estimate maximum possible file size needed
    i = 1.3 * i + 1000 + Len(m_Comment)
    ReDim m_Data(i)
    m_Ptr = 0

    InsertMarker SOI                                   'SOI - Start of Image
    InsertJFIF                                         'JFIF

    If Len(m_Comment) > 0 Then InsertCOM m_Comment     'COM - Comment
    InsertCOM "Saved by PhotoDemon v" & App.Major & "." & App.Minor & " using John Korejwa's JPEG encoder (v0.9). "

    InsertDQT m_Ptr, 0                                 'DQT - Define Quantization Tables
    If Nf > 1 Then InsertDQT m_Ptr, 1

    InsertSOF SOF0                                     'SOF - Start of Frame

    InsertDHT m_Ptr, 0, False                          'DHT - Define Huffman Tables
    InsertDHT m_Ptr, 0, True
    If Nf > 1 Then
        InsertDHT m_Ptr, 1, False
        InsertDHT m_Ptr, 1, True
    End If

    InsertSequentialScans CompIndex, Td, Ta, 0, Nf - 1 'SOS - Scan Data
    InsertMarker EOI                                   'EOI - End of Image

    Message "Writing data to file..."

   'Size the final byte array and write to file
    ReDim Preserve m_Data(m_Ptr - 1)
    fileNum = FreeFile
    Open FileName For Binary Access Write As fileNum
        Put #fileNum, , m_Data
    Close fileNum
    Erase m_Data


        End If
    End If

End Function




'========================================================================================
'                         C L A S S   I N I T I A L I Z A T I O N
'========================================================================================
Private Sub Class_Initialize()
    Dim i As Long
    Dim j As Long
    Dim dx As Long
    Dim ZZ As Long

    i = 0                   'Initialize the ZigZag() array, which maps out the
    j = 0                   '  zig-zag sequence of quantized DCT coefficients
    dx = 1                  '  in approximately low to high spatial frequencies
    For ZZ = 0 To 63
        ZigZag(i, j) = ZZ
        i = i + dx
        j = j - dx
        If i > 7 Then              '  0   1   5   6  14  15  27  28
            i = 7                  '  2   4   7  13  16  26  29  42
            j = j + 2              '  3   8  12  17  25  30  41  43
            dx = -1                '  9  11  18  24  31  40  44  53
        ElseIf j > 7 Then          ' 10  19  23  32  39  45  52  54
            j = 7                  ' 20  22  33  38  46  51  55  60
            i = i + 2              ' 21  34  37  47  50  56  59  61
            dx = 1                 ' 35  36  48  49  57  58  62  63
        ElseIf i < 0 Then
            i = 0 'check (j>7) first
            dx = 1
        ElseIf j < 0 Then
            j = 0
            dx = -1
        End If
    Next ZZ

   'Luminance Quantization table for Quality = 50
    QLumin(0) = 16:   QLumin(1) = 11:   QLumin(2) = 12:   QLumin(3) = 14
    QLumin(4) = 12:   QLumin(5) = 10:   QLumin(6) = 16:   QLumin(7) = 14
    QLumin(8) = 13:   QLumin(9) = 14:   QLumin(10) = 18:  QLumin(11) = 17
    QLumin(12) = 16:  QLumin(13) = 19:  QLumin(14) = 24:  QLumin(15) = 40
    QLumin(16) = 26:  QLumin(17) = 24:  QLumin(18) = 22:  QLumin(19) = 22
    QLumin(20) = 24:  QLumin(21) = 49:  QLumin(22) = 35:  QLumin(23) = 37
    QLumin(24) = 29:  QLumin(25) = 40:  QLumin(26) = 58:  QLumin(27) = 51
    QLumin(28) = 61:  QLumin(29) = 60:  QLumin(30) = 57:  QLumin(31) = 51
    QLumin(32) = 56:  QLumin(33) = 55:  QLumin(34) = 64:  QLumin(35) = 72
    QLumin(36) = 92:  QLumin(37) = 78:  QLumin(38) = 64:  QLumin(39) = 68
    QLumin(40) = 87:  QLumin(41) = 69:  QLumin(42) = 55:  QLumin(43) = 56
    QLumin(44) = 80:  QLumin(45) = 109: QLumin(46) = 81:  QLumin(47) = 87
    QLumin(48) = 95:  QLumin(49) = 98:  QLumin(50) = 103: QLumin(51) = 104
    QLumin(52) = 103: QLumin(53) = 62:  QLumin(54) = 77:  QLumin(55) = 113
    QLumin(56) = 121: QLumin(57) = 112: QLumin(58) = 100: QLumin(59) = 120
    QLumin(60) = 92:  QLumin(61) = 101: QLumin(62) = 103: QLumin(63) = 99

   'Chrominance Quantization table for Quality = 50
    QChrom(0) = 17:   QChrom(1) = 18:   QChrom(2) = 18:   QChrom(3) = 24
    QChrom(4) = 21:   QChrom(5) = 24:   QChrom(6) = 47:   QChrom(7) = 26
    QChrom(8) = 26:   QChrom(9) = 47:   QChrom(10) = 99:  QChrom(11) = 66
    QChrom(12) = 56:  QChrom(13) = 66:  QChrom(14) = 99:  QChrom(15) = 99
    QChrom(16) = 99:  QChrom(17) = 99:  QChrom(18) = 99:  QChrom(19) = 99
    QChrom(20) = 99:  QChrom(21) = 99:  QChrom(22) = 99:  QChrom(23) = 99
    QChrom(24) = 99:  QChrom(25) = 99:  QChrom(26) = 99:  QChrom(27) = 99
    QChrom(28) = 99:  QChrom(29) = 99:  QChrom(30) = 99:  QChrom(31) = 99
    QChrom(32) = 99:  QChrom(33) = 99:  QChrom(34) = 99:  QChrom(35) = 99
    QChrom(36) = 99:  QChrom(37) = 99:  QChrom(38) = 99:  QChrom(39) = 99
    QChrom(40) = 99:  QChrom(41) = 99:  QChrom(42) = 99:  QChrom(43) = 99
    QChrom(44) = 99:  QChrom(45) = 99:  QChrom(46) = 99:  QChrom(47) = 99
    QChrom(48) = 99:  QChrom(49) = 99:  QChrom(50) = 99:  QChrom(51) = 99
    QChrom(52) = 99:  QChrom(53) = 99:  QChrom(54) = 99:  QChrom(55) = 99
    QChrom(56) = 99:  QChrom(57) = 99:  QChrom(58) = 99:  QChrom(59) = 99
    QChrom(60) = 99:  QChrom(61) = 99:  QChrom(62) = 99:  QChrom(63) = 99


    FDCTScale(0) = 0.353553390593273     '0.25 / Cos(4 / 16 * PI)
    FDCTScale(1) = 0.25489778955208      '0.25 / Cos(1 / 16 * PI)
    FDCTScale(2) = 0.270598050073098     '0.25 / Cos(2 / 16 * PI)
    FDCTScale(3) = 0.300672443467523     '0.25 / Cos(3 / 16 * PI)
    FDCTScale(4) = 0.353553390593273     '0.25 / Cos(4 / 16 * PI)
    FDCTScale(5) = 0.449988111568207     '0.25 / Cos(5 / 16 * PI)
    FDCTScale(6) = 0.653281482438186     '0.25 / Cos(6 / 16 * PI)
    FDCTScale(7) = 1.28145772387074      '0.25 / Cos(7 / 16 * PI)

    SetSamplingFrequencies 2, 2, 1, 1, 1, 1
    Quality = 75

End Sub
