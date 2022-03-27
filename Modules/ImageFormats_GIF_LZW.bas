Attribute VB_Name = "ImageFormats_GIF_LZW"
'***************************************************************************
'LZW encoder for GIF export
'Derived from work originally copyright by multiple authors; see below for licensing details
'Created: 13/October/21
'Last updated: 27/March/22
'Last update: enable LZW heuristics for improved compression ratios
'
'In 2021, I spent some time optimizing PhotoDemon's animated GIF exporter.  Many optimizations
' occur before actual GIF encoding (frame- and palette-related stuff), but for final GIF encoding
' I leaned on the 3rd-party FreeImage library.  Unfortunately, FreeImage's GIF support is mediocre
' (palettes are always written as 256-color tables, encoding is built atop strings (!!!) so perf
' is rough), so I started hunting for alternative solutions.
'
'After getting angry at giflib and the incredibly unpleasant shitshow that is compiling it as
' a Windows DLL, I stumbled across a VB6 LZW encoder at the planet-source-code archive on GitHub:
'
'https://github.com/Planet-Source-Code/carles-p-v-image-8-bpp-ditherer-native-gif-encoder__1-45899
'
'Carles's project is a modified version of a VB6 LZW implementation originally by Ron van Tilburg:
'
'https://github.com/Planet-Source-Code/ron-van-tilburg-rvtvbimg__1-14210/blob/master/GIFSave.bas
'
'Per the comments in his code, Ron translated his initial version from C sources derived from the
' original UNIX compress.c.  It was a fun trip down memory lane to see some familiar PSC authors,
' but as always, there's a critical problem with VB6 code from public archives like this -
' the likelihood of the code being exhaustively stress-tested (for performance, security,
' reliability, etc) is... not so likely.
'
'Fortunately, the original C sources for compress.c are still available, as are translations into
' myriad other languages - for example, a few seconds on GitHub turned up these:
' https://github.com/mlapadula/GifEncoder/blob/master/lib/src/main/java/com/mlapadula/gifencoder/LzwEncoder.java
' https://github.com/P1ayer4312/Steam-Artwork-Cropper/blob/main/steam/js/gif.js-master/src/LZWEncoder.js
'
'Armed with multiple references to compare and contrast, I set about mixing and matching the
' original C code with some ideas from Carles/Ron's VB6 versions to produce something appropriate
' for PhotoDemon.  I think the final result is very good, with meaningfully improved performance,
' a number of fixed edge-case bugs, reworking of various suboptimal-for-VB6 designs, and improved
' LZW encoding efficiency.  The final result is a very compact LZW encoder with efficiency and
' performance on par with giflib, and pretty much on-par with the original compress.c version.
'
'Note that this module only handles the LZW encoding portion of GIF export.  All the actual file
' encoding (including headers, tables, etc) is 100% my own work and it remains in the
' ImageFormats_GIF module under the same BSD license as PD itself.  I've deliberately moved this
' LZW code into its own file because the 40-year copyright list is quite long, and licensing is
' murky because while compress.c was released into the public domain, the licensing state of
' subsequent translations is unclear.  For commercial usage you may want to contact some of the
' authors listed below for clarity, but please consider any modifications I've made to be freely
' licensed in the public domain under the unlicense (https://unlicense.org/).
'
'Anyway, here's my attempt to merge various credit lists from the sources mentioned above into
' one comprehensive list:
'
'GIF Image compression based on: compress.c (File compression ala IEEE Computer, June 1984.)
' Original Authors:             Spencer W. Thomas       (decvax!harpo!utah-cs!utah-gr!thomas)
'                               Jim McKie               (decvax!mcvax!jim)
'                               Steve Davies            (decvax!vax135!petsd!peora!srd)
'                               Ken Turkowski           (decvax!decwrl!turtlevax!ken)
'                               James A. Woods          (decvax!ihnp4!ames!jaw)
'                               Joe Orost               (decvax!vax135!petsd!joe)
'                               David Rowley            (mgardi@watdcsu.waterloo.edu)
' Initial VB6 translation by:   Ron van Tilburg         (rivit@f1.net.au)
' ...with additional VB6-specific modifications by Carles P.V. and Tanner Helland (tannerhelland.com)
'
'***************************************************************************

Option Explicit

'GIF's LZW variant has specific requirements.  12-bits is the longest allowable LZW code,
' which limits the code dictionary to [0, 4095].  Other constants derive from that.
Private Const MAX_BITS As Long = 12             'From the GIF spec (https://www.w3.org/Graphics/GIF/spec-gif89a.txt)
Private Const MAX_BITSHIFT As Long = 4096       '2 ^ MAX_BITS, used to shift variable-length bit codes
Private Const MAX_CODE_DEFAULT As Long = 4096   'Size of LZW dict for a compressed GIF (uncompressed works differently)
Private MAX_CODE As Long                        'Will be MAX_CODE_DEFAULT unless *uncompressed* option is selected;
                                                ' (uncompressed depends on original code size, e.g. palette size)
Private Const EOF_CODE As Long = -1             'Internal marker for EOF
Private Const TABLE_SIZE As Long = 5003         'Hash table needs prime size, aim for 80% occupancy (80% of 5003 = ~4096)

'To enable compression heuristics (which attempt to clear the LZW code table at "ideal" intervals),
' set this value >= 0.  The current design requires the value to be of the form (2^n - 1), because it
' uses AND masking to mark intervals.
Private Const ENABLE_HEURISTICS As Long = 255

'Same as ENABLE_HEURISTICS, above, but specifically designed for the special case of
' "after the LZW code table fills".  The way the GIF format is designed, you only get 4096 unique
' LZW values before you need to clear the table and start over - but you don't *have* to start over
' at that point.  You can continue using the current table as-is, without adding new entries to it.
' Because resetting the table also resets your current compression ratio, each block has a
' "sweet spot" where the current table is still doing better work than a reset would provide -
' but you have to test for this state using heuristics, because the "sweet spot" varies wildly from
' block to block.
'
'When set to 0, the table always clears immediately after it reaches capacity, (which is the behavior
' most GIF encoders use).
Private Const ENABLE_HEURISTICS_FULL As Long = 31

'Name says it all: starts at 9 for 8-bpp data (because 8 bits are reserved for standard values [0-255])
' and increases as table fills.  Smaller palettes start at the nearest power-of-two of their palette size,
' with some exceptions (1-bit images are manually up-sized a bit, for example, per the GIF spec).
Private m_bitsPerCode As Long

'Max code value, given m_bitsPerCode (works as an index into the m_Masks array)
Private m_maxCode As Long
Private m_masks(0 To 16) As Long    '(2 ^ n) - 1, see InitMasks() function

'Hash table stores indices into code table.  This is an implementation detail only;
' LZW can be implemented any number of other ways.
Private m_hashTable(0 To TABLE_SIZE - 1) As Long
Private m_codeTable(0 To TABLE_SIZE - 1) As Long

'A special "Clear code" is always defined as 2 ^ bits-per-pixel.  So for a 256-color image
' (8-bpp), the clear code is exactly 256.  From the GIF spec:
'
' "A special Clear code is defined which resets all compression/decompression parameters and
' tables to a start-up state... The Clear code can appear at any point in the image data stream
' and therefore requires the LZW algorithm to process succeeding codes as if a new data stream
' was starting. Encoders should output a Clear code as the first code of each image data stream."
'
'Declared in all caps because it's set once at encoding start, then subsequently treated as
' a constant (contingent on the initial code size).
Private LZW_CLEAR_CODE As Long

'EOF code is always ClearCode + 1.  From the GIF spec:
' "An End of Information code is defined that explicitly indicates the end of the image data
' stream. LZW processing terminates when this code is encountered. It must be the last code output
' by the encoder for an image."
'
'As with the clear code, declared in all caps because it's set once at encoding start, then
' subsequently treated as a constant.
Private LZW_EOF_CODE As Long

'First unused entry (index) in the table.  Starts as ClearCode + 2 and resets to that value
' whenever the table is reset (typically when the table fills, although you can technically
' continue writing out values without actually storing them in the table).
Private m_freeEntry As Long

'A major limitation of ZLW encoding is the fixed maximum size of the code table.  When the table
' is filled, there's no way to remove elements from it - you have to flush the entire table and
' start over from scratch.  (Contrast this with LZ77's "sliding window" which doesn't use a table;
' instead it uses references to previous points in the stream, which means data automatically
' "falls out of scope" as the window slides across the data.)
'
'When a table clear needs to be initiated, we set a flag that tells several functions to modify
' their behavior accordingly.  (Many values need to be reset along with the table, like the
' initial bit-size of codes.)
Private m_tableJustCleared As Boolean

'Initial length of output codes varies according to the underlying palette size.  8-bit palettes
' must start at 9-bits, but smaller palettes are allowed to start at a smaller level (and thus
' have more room in their table before filling it).  From the GIF spec:
'
' "The output codes are of variable length, starting at <code size>+1 bits per code, up to 12 bits
' per code. This defines a maximum code value of 4095 (0xFFF). Whenever the LZW code value would
' exceed the current code length, the code length is increased by one. The packing/unpacking of
' these codes must then be altered to reflect the new code length."
Private m_initBits As Long

'No local copies are made of the original pixel data.  Instead, we (unsafely) wrap an array
' around the source data pointer.  These indices track the current source pixel and the maximum
' source pixel (used to detect completion).
Private m_curPixel As Long, m_totalPixels As Long

'A simple 32-bit int works well for accumulating bits.  Once 8+ bits have been added, we flush
' them out to the current black 1-byte at a time.  Remaining bits are then shifted over, and
' prebuilt masks (m_Masks, above) are used to add the next set of bits to the "bucket".
'
'm_codeBits tracks how many code bits are currently in the bucket.  (This is used as an
' index into the mask array, among other things.)
Private m_codeBucket As Long, m_codeBits As Long

'GIF's LZW implementation requires you to flush collected codes out to file in 255-byte
' increments (max size; blocks can be smaller if you want).  The final LZW stream is thus
' a series of 1-byte block size indicators followed by [n] bytes of LZW-encoded data.
Private m_curBlock(0 To 255) As Byte, m_blockSize As Long

'PD-specific objects:

'PhotoDemon wraps this array around a source unsigned char array (palettized pixel data)...
Private m_srcData() As Byte

'...and directly dumps encoded output to a pdStream object (which is a memory-mapped interface
' around the target GIF file)
Private m_dstStream As pdStream

'We use heuristics to improve encoding efficiency.  If efficiency starts trending downward
' (we check it on every sub-block write, and look for two consecutive "worsening" blocks),
' it's almost always better to just clear the LZW table and start over at a smaller code size.
' This simple strategy produces significantly better results on extremely noisy images -
' particularly GIFs from video sources - because LZW requires ever-increasing code sizes,
' which produces a disproportionately high penalty relative to just writing uncompressed data
' when the input data is noisy.
Private m_BytesOut As Long, m_BytesIn As Long
Private m_CurRatio As Single, m_LastRatio As Single, m_NumWorseningBlocks As Long
Private Const LAST_RATIO_RESET As Single = -1!

'When writing uncompressed GIFs, turn all heuristics off
Private m_writeUncompressed As Boolean

'This table serves double duty; masks are used to mask off increasing bit lengths (2 ^ n-1),
' but we can also add 1 to a given mask to get a power-of-two to help work around VB's lack
' of shift operators.
Private Sub CacheMasks()
    m_masks(0) = &H0&
    m_masks(1) = &H1&
    m_masks(2) = &H3&
    m_masks(3) = &H7&
    m_masks(4) = &HF&
    m_masks(5) = &H1F&
    m_masks(6) = &H3F&
    m_masks(7) = &H7F&
    m_masks(8) = &HFF&
    m_masks(9) = &H1FF&
    m_masks(10) = &H3FF&
    m_masks(11) = &H7FF&
    m_masks(12) = &HFFF&
    m_masks(13) = &H1FFF&
    m_masks(14) = &H3FFF&
    m_masks(15) = &H7FFF&
    m_masks(16) = &HFFFF&
End Sub

'Main LZW encode function.  Outputs a GIF-compatible LZW stream to dstStream, encoding the data
' from ptrToSourceImage. lzwCodeSize is allowed to float (but must be at least *2*); GIFs set this
' according to the palette size, with shorter code sizes for smaller palettes.
Public Sub CompressLZW(ByRef dstStream As pdStream, ByVal ptrToSourceImage As Long, ByVal sizeOfSourceImage As Long, ByVal lzwCodeSize As Byte, Optional ByVal useCompression As Boolean = True)
    
    'Ensure bit-masks are available
    CacheMasks
    
    'Create a blank stream object.  We'll dump our LZW results to memory, and if they compress
    ' well we'll dump them from there to file.  (But if they compress poorly, we'll dump
    ' uncompressed bytes instead - on noisy inputs, GIF's LZW scheme can produce significantly
    ' larger results with compression than without.)
    Set m_dstStream = New pdStream
    m_dstStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, startingBufferSize:=sizeOfSourceImage
    
    'Set a local reference to source data (image bytes, already palettized)
    Dim tmpSA1D As SafeArray1D
    VBHacks.WrapArrayAroundPtr_Byte m_srcData, tmpSA1D, ptrToSourceImage, sizeOfSourceImage
    
    'Set initial and final positions
    m_curPixel = 0
    m_totalPixels = sizeOfSourceImage
    
    'If *not* using compression, specify a much smaller max size (which will trigger early table resets,
    ' but skip any actual LZW processing).
    m_writeUncompressed = (Not useCompression)
    If m_writeUncompressed Then
        
        'For uncompressed GIFs (which actually produce better results on extremely noisy images), use a shorter
        ' max code that eliminates LZW behavior; see https://en.wikipedia.org/wiki/GIF#Uncompressed_GIF for details
        MAX_CODE = (2 ^ lzwCodeSize) - 2
        
    Else
        MAX_CODE = MAX_CODE_DEFAULT
    End If
    
    'Note initial code size (encoder needs to know it to prep the initial code table),
    ' then launch LZW encoder.
    m_initBits = lzwCodeSize
    CompressAndWriteBits
    
    'If compression is being used, ensure that it actually produced smaller results.
    If useCompression Then
    
        'Calculate a (reasonably) good estimate of the size of an uncompressed GIF stream.
        ' This is usually the input size * (9/8) (if 8-bits are used, [lzwcodesize]/[lzwcodesize-1] otherwise) +
        ' input size * (1 / 256) for block markers.
        Dim expectedSize As Long
        expectedSize = Int(sizeOfSourceImage * CDbl(lzwCodeSize / (lzwCodeSize - 1)))
        expectedSize = expectedSize + Int(sizeOfSourceImage * CDbl(1# / 256#))
        
        'Since this is just an estimate, add a 1% overhead to ensure we don't switch to
        ' uncompressed output prematurely.
        expectedSize = Int(expectedSize * 1.01)
        
        'If our compressed stream exceeded that size, dump it out uncompressed instead
        If (m_dstStream.GetStreamSize > expectedSize) Then
            
            PDDebug.LogAction "Note: compressed GIF is larger than uncompressed version; writing uncompressed version instead."
            PDDebug.LogAction "(" & m_dstStream.GetStreamSize & " vs " & expectedSize & ")"
            
            'Reset all initial inputs and start over.
            Set m_dstStream = New pdStream
            m_dstStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, startingBufferSize:=expectedSize
            
            m_curPixel = 0
            m_totalPixels = sizeOfSourceImage
            
            m_writeUncompressed = True
            MAX_CODE = (2 ^ lzwCodeSize) - 2
            m_initBits = lzwCodeSize
            
            'Take two!
            CompressAndWriteBits
            
        End If
    
    End If
    
    'Dump our temporary stream to the destination stream and remove unsafe array wrapper before exiting.
    dstStream.WriteStream m_dstStream
    VBHacks.UnwrapArrayFromPtr_Byte m_srcData
    
End Sub

'LZW encoder.  Only safe to call from Encode(), above, due to various required initialization steps.
Private Sub CompressAndWriteBits()
    
    'In some ways, LZW is basically just a process of moving data into and out of a series of tables.
    ' This generic index is used to index into code and hash tables, primarily.
    Dim idxTable As Long
    idxTable = 0
    
    'Start with an empty bucket at bit-offset 0
    m_codeBucket = 0
    m_codeBits = 0
    
    'Establish some "constants" given the initial bit-size of LZW codes for this image
    m_bitsPerCode = m_initBits
    m_maxCode = m_masks(m_bitsPerCode)
    LZW_CLEAR_CODE = 2 ^ (m_initBits - 1)
    LZW_EOF_CODE = LZW_CLEAR_CODE + 1
    m_freeEntry = LZW_CLEAR_CODE + 2
    
    'Initialize any remaining flags and/or trackers
    m_tableJustCleared = False
    m_blockSize = 0
    VBHacks.ZeroMemory VarPtr(m_curBlock(0)), 256
    
    m_LastRatio = LAST_RATIO_RESET
    m_NumWorseningBlocks = 0
    m_BytesIn = 0
    m_BytesOut = 0
    
    'Prep hash table
    Dim hShift As Long, fCode As Long
    hShift = 0
    fCode = TABLE_SIZE
    Do While (fCode < 65536)
        hShift = hShift + 1
        fCode = fCode * 2
    Loop
    
    'Set hash code range bound
    hShift = m_masks(8 - hShift) + 1
    
    'Initialize a default hash table (initialized to "-1", because 0 is a valid indicator for "empty")
    ClearTable
    
    'From the GIF spec:
    ' "Encoders should output a Clear code as the first code of each image data stream."
    OutputCode LZW_CLEAR_CODE
    
    'Retrieve the first two pixels, then get the show underway!
    Dim lEnt As Long
    lEnt = GetNextPixel()
    
    Dim lC As Long
    lC = GetNextPixel()
    
    Do While (lC <> EOF_CODE)
        
        'Calculate index into the hash table
        fCode = lC * MAX_BITSHIFT + lEnt
        idxTable = (lC * hShift) Xor lEnt
        
        'If the hash table provides an immediate hit for the current pattern, return it and move on
        If (m_hashTable(idxTable) = fCode) Then
            lEnt = m_codeTable(idxTable)
            GoTo NextPixel
        
        'If the hash table hits an empty index, we'll add this code to the table *then* move on.
        ' (The hash table is initialized to -1.)
        ElseIf (m_hashTable(idxTable) < 0) Then
            GoTo NoMatch
        End If
        
        'The hash table hit an entry, but not one matching the current pattern.
        ' Hash it again and see what we find.
        Dim lDisp As Long
        lDisp = TABLE_SIZE - idxTable
        If (idxTable = 0) Then lDisp = 1    'Failsafe to ensure we change position

        '(As long as we find entries in the table but not the one we want, we'll keep jumping
        ' around the table until we hit and/or miss)
Probe:
        idxTable = idxTable - lDisp
        If (idxTable < 0) Then idxTable = idxTable + TABLE_SIZE
        
        'If we found a match, return it and carry on
        If (m_hashTable(idxTable) = fCode) Then
            lEnt = m_codeTable(idxTable)
            GoTo NextPixel
        End If
        
        'If we found an entry but *not* the one we want, carry on
        If (m_hashTable(idxTable) >= 0) Then GoTo Probe

        'If we're here, we found an empty entry in the hash table.  Let's store our current code.
NoMatch:
        
        'Write out our current code
        OutputCode lEnt
        lEnt = lC
        
        'If we have room in the table, add this entry to both the code table
        ' *and* the hash table.
        If (m_freeEntry < MAX_CODE) Then
        
            m_codeTable(idxTable) = m_freeEntry
            m_freeEntry = m_freeEntry + 1  'Code -> Hash table
            m_hashTable(idxTable) = fCode
            
            'Every (x) bytes of input, run some heuristics to see if it's worthwhile to clear
            ' the LZW table and start over.
            If (Not m_writeUncompressed) Then
                
                'Ignore compression statistics until we start reaching byte-for-byte rounds of LZW work.
                ' (Compression tends to be extremely noisy until bit-width reaches ~10 bits, and it's easy
                ' to get caught in a bad local minima that looks like worsening compression when actually
                ' it's just priming the table for later successes.)
                If (m_bitsPerCode > 9) Then
                    
                    'Only check every [ENABLE_HEURISTICS] number of input bytes
                    If (ENABLE_HEURISTICS And m_BytesIn) Then
                    
                        'Because we're dealing with ratios here (not absolute values), minor differences
                        ' can signal larger problems, especially when jumping up a bit-count.  Look for
                        ' 3 (2+1) successive drops of 0.5% in our overall ratio; if this happens,
                        ' compression rates are dropping quickly and we could possibly benefit from a
                        ' fresh table.
                        AssessCompressionHeuristics 2, 0.995!
                        
                    End If
                    
                End If
                
            End If
            
        'If we don't have room, reset the table
        Else
            
            'To never reset the table, do nothing here!
            
            'To reset the table immediately after it fills, perform a clear operation here:
            If m_writeUncompressed Then
                ProcessFullTableClear
            Else
                
                'Perhaps you want to skip a reset if you're close to the end of the frame, since you
                ' won't have time to "build up" meaningful compression results again - so try something
                ' like this:
                If ((m_totalPixels - m_curPixel) < 256) Then
                    'Keep using the table as-is since it won't hurt anything this close to the end
                Else
                
                    'Finally, if you want to do meaningful heuristics, call the heuristics assessor
                    ' with whatever pixel threshold you want (but don't check on every pixel -
                    ' that's wasteful!)  In PD, we check every 32 bytes past the end of the table
                    ' and as soon as we detect a compression ratio drop, we dump the table.
                    If (ENABLE_HEURISTICS_FULL And m_BytesIn) Then
                        AssessCompressionHeuristics 0, 1!
                    
                    'Ensure we have a good "last ratio" value now that the table is full
                    ElseIf (m_LastRatio = LAST_RATIO_RESET) Then
                        AssessCompressionHeuristics 0, 1!
                    End If
                
                End If
                
            End If
            
        End If
        
        'Work complete!  Grab the next pixel and carry on
NextPixel:
        lC = GetNextPixel()
    
    '/do while (lC <> EOF_CODE)
    Loop

    'Output the final code, then mark EOF and exit
    OutputCode lEnt
    OutputCode LZW_EOF_CODE
    
End Sub

'Freely modify (and call!) this function at your leisure.
'
'This function was added by Tanner Helland (https://github.com/tannerhelland) for exploring dynamic
' LZW code table resets (in hopes of producing smaller GIF foiles).  There's not a lot of good
' scholarly research on this topic, but for an excellent overview of the general idea, see here:
'
' https://www.nayuki.io/page/gif-optimizer-java
'
'As for clearing the LZW table, nearly all GIF encoders take one of two approaches:
' 1) Clear the LZW table when it reaches maximum size, or...
' 2) Clear the LZW table 2 spaces before it reaches its first code size increase
'    (which produces a so-called "uncompressed" GIF - which despite the name, actually produces
'     smaller results than LZW itself when the input frame is extremely noisy)
'
'A "perfect" GIF encoder would actually work somewhere in-between these two mechanisms, filling the
' LZW table as long as it's helping compression, then clearing it whenever compression drops.
' Note that you also don't *have* to reset the code table when it reaches max size - in fact, on many
' images you may be able to eke out small gains by continuing to use the full code table as long as
' it produces "good" results.  (This is a quirk of LZW encoding, since a table clear also resets the
' compression ratio back to 1:1 until the table "jumps" a code size and starts beating out bare inputs
' again.)
'
'Anyway, the function below works fine, but I haven't cracked the logic necessary to really perfect
' a clearing strategy.  I've found it very hard to address outliers in particular - extremely
' compressible or extremely incompressible inputs - without hurting more normal images.  More work
' is needed!
'
'Returns TRUE if table was cleared by heuristics; FALSE otherwise.
Private Function AssessCompressionHeuristics(ByVal consecutiveMissesAllowed As Long, Optional ByVal lastRatioModifier As Single = 1!) As Boolean

    'Whenever you want to check the current compression ratio, do it here.  Just remember to reset
    ' the bytes in/out values after each table clear.
    If (m_BytesOut <> 0) Then
        
        'Calculate the current encoding ratio
        m_CurRatio = m_BytesIn / m_BytesOut
        
        'If we have a previous ratio, compare the two.
        If (m_LastRatio > 0!) Then
            
            'If this ratio is worse than our previous ratio by some threshold (currently 1%), note it.
            ' We are especially interested in successively worse compression ratios, but the critical
            ' question is - how frequently to check said intervals, and what threshold determines when
            ' to clear the table?
            '
            'A good solution would likely involve statistical modeling (hidden Markov models comes to mind),
            ' but a naive solution can also provide some benefit, I hope...
            If (m_CurRatio < (m_LastRatio * lastRatioModifier)) Then
                m_NumWorseningBlocks = m_NumWorseningBlocks + 1
                    
                    'Three bad checks in a row should mean it's okay to dump the current table
                    If (m_NumWorseningBlocks > consecutiveMissesAllowed) Then
                        ProcessFullTableClear
                        AssessCompressionHeuristics = True
                    End If
                    
            'Reset the "consecutive worsening blocks" tracker, since compression has improved
            ' since our last check.
            Else
                m_NumWorseningBlocks = 0
            End If
            
            m_LastRatio = m_CurRatio
            
        'If we don't have a previous ratio, the table was just reset.  Note the current ratio,
        ' and we'll compare it against the *next* one.
        Else
            m_LastRatio = m_CurRatio
        End If
    
    End If
    
End Function

Private Sub ProcessFullTableClear()

    'After clearing the table, we need to reset the pointer into the table
    ClearTable
    m_freeEntry = LZW_CLEAR_CODE + 2
    m_tableJustCleared = True
    
    'A clear code also needs to be placed in the stream, so the decompressor
    ' knows to reset its table too.  (In LZW encoding, tables are never stored.  They are
    ' reconstructed on-the-fly by the decoder, and clear codes ensure that the table used
    ' by the encoder can be correctly synced by the decoder.)
    OutputCode LZW_CLEAR_CODE
    
End Sub

'Safe wrapper for pixel retrieval; when EOF is reached, return a special (internal) end code
Private Function GetNextPixel() As Long
    
    If (m_curPixel < m_totalPixels) Then
        GetNextPixel = m_srcData(m_curPixel)
        m_curPixel = m_curPixel + 1
    Else
        GetNextPixel = EOF_CODE
    End If
    
    'Only required for compression heuristics (you can safely ignore otherwise)
    m_BytesIn = m_BytesIn + 1
    
End Function

'Add the passed variable-length code to the output bucket.  If we have more than 8-bits of completed data,
' 1+ filled bytes will be copied out to the current block.
Private Sub OutputCode(ByVal newCode As Long)
    
    'Mask off any remaining bits that have been flushed out to file.
    m_codeBucket = (m_codeBucket And m_masks(m_codeBits))
    
    'If we already have data in the bucket, shift it before adding the new code
    If (m_codeBits > 0) Then
        m_codeBucket = m_codeBucket Or (newCode * (1 + m_masks(m_codeBits)))
    
    '...otherwise, just dump the new code into the bucket
    Else
        m_codeBucket = newCode
    End If
    
    'Remember how many bits were just added to the bucket
    m_codeBits = m_codeBits + m_bitsPerCode

    'If we have at least 1 full byte worth of code bits, dump them out to the current block,
    ' then shift the stored codes over 8-bits (and note that we just lost 8-bits worth of data)
    Do While (m_codeBits >= 8)
    
        AddCharToBlock m_codeBucket And &HFF&
        m_codeBucket = m_codeBucket \ 256&
        m_codeBits = m_codeBits - 8
        
        'While here, track bytes "written".
        ' (This is only used for compression heuristics; it is not part of standard LZW compression.)
        m_BytesOut = m_BytesOut + 1
        
    Loop
    
    'Before exiting, perform some house-keeping.
    
    'If our parent just cleared the code table, we need to reset our internal bit-count to match.
    ' (Bits-per-code increases as the table fills.  When the table is flushed, we reset to
    '  9-bits-per-code for 8-bpp data, fewer bits for smaller palettes).
    If m_tableJustCleared Then
        m_bitsPerCode = m_initBits
        m_maxCode = m_masks(m_bitsPerCode)
        m_tableJustCleared = False
    
    'If the table wasn't cleared but we just used up the last entry in the current bit-size,
    ' we need to increase the current code bit-size to match
    ElseIf (m_freeEntry > m_maxCode) Then
        
        m_bitsPerCode = m_bitsPerCode + 1
        
        'We've maxed out the table.  Increase bits-per-code.  (Note that once we hit 12-bit codes,
        ' we stay there.  Our parent may not choose to reset the table just yet if we're still
        ' achieving a good compression ratio.)
        If (m_bitsPerCode = MAX_BITS) Then
            m_maxCode = MAX_CODE
        Else
            m_maxCode = m_masks(m_bitsPerCode)
        End If
        
    End If
    
    'One last bit of house-keeping.  If we were just passed the special EOF code,
    ' flush out everything we have in the buffer.
    If (newCode = LZW_EOF_CODE) Then
    
        Do While (m_codeBits > 0)
            AddCharToBlock m_codeBucket And &HFF&
            m_codeBucket = m_codeBucket \ 256&
            m_codeBits = m_codeBits - 8
        Loop
        
        'Before exiting, ensure one final flush out to file
        FlushBlock
        
    End If
    
End Sub

'Reset the hash table.  -1 = "empty" since 0 is a valid index.
' FillMemory is used to flush the table quickly.
Private Sub ClearTable()
    
    VBHacks.FillMemory VarPtr(m_hashTable(0)), (UBound(m_hashTable) + 1) * 4, &HFF&
    
    'Reset all heuristics, since the table is now starting over from "0"
    m_BytesIn = 0
    m_BytesOut = 0
    m_LastRatio = LAST_RATIO_RESET
    m_NumWorseningBlocks = 0
    
End Sub

'Add a char to the current packet and flush when max capacity is reached
Private Sub AddCharToBlock(ByVal lChar As Long)
    m_blockSize = m_blockSize + 1       'Start at position [1], leaving position [0] for block size indicator
    m_curBlock(m_blockSize) = lChar
    If (m_blockSize = 255) Then FlushBlock
End Sub

'Place the current packet (up to 255 entries) to the destination stream, and reset the
' accumulator table.
Private Sub FlushBlock()
    
    'Don't flush an empty block (0 block size is special EOF marker)
    If (m_blockSize > 0) Then
        
        'Write block length (single byte, max 255) followed by block itself
        m_curBlock(0) = m_blockSize
        m_dstStream.WriteBytesFromPointer VarPtr(m_curBlock(0)), m_blockSize + 1
        
        'Reset the accumulator and note that we don't need to zero the table;
        ' it's automatically overwritten as necessary.
        m_blockSize = 0
        
    End If
    
End Sub
