VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdHGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Shuttle Radar Topography Mission (SRTM)" (HGT) Decoder
'Copyright 2022-2026 by Tanner Helland
'Created: 13/September/22
'Last updated: 15/November/22
'Last update: harden against malicious file input
'
'The HGT file format contains a raw dump of satellite topography data.  GIMP added support for this
' format back in 2017 (https://girinstud.io/news/2017/12/new-format-in-gimp-hgt/) and I thought it
' might be fun to do the same in PD.
'
'You can find the HGT spec here (link good as of Sep 2022):
'
' https://github.com/wellstorm/srtm/blob/master/SRTM_Topo.txt
'
'My preferred interface for downloading topography data for anywhere on planet earth is this cool site:
'
' https://dwtkns.com/srtm30m/
'
'You will need a free NASA Earthdata login; a link to sign-up is provided at that site.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this dumps additional diagnostic information
' to PD's primary debug log.
Private Const HGT_DEBUG_VERBOSE As Boolean = False

'Image width/height, in pixels (big-endian)
Private m_Width As Long, m_Height As Long

'Byte-by-byte access is provided, as always, by a pdStream instance.
Private m_Stream As pdStream

'HGT files are raw data dumps, so they do not provide an easy validation mechanism - but because they
' contain fixed tile counts, we can perform a rough validation using file size.
Friend Function IsFileHGT(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True) As Boolean
    
    Dim potentialMatch As Boolean
    potentialMatch = Files.FileExists(srcFilename)
    
    'Per the spec...
    ' "International 3-arc-second files have 1201 columns and 1201 rows of data, with a total filesize
    '  of 2,884,802 bytes ( = 1201 x 1201 x 2). United States 1-arc-second files have 3601 columns and
    '  3601 rows of data, with a total filesize of 25,934,402 bytes ( = 3601 x 3601 x 2)."
    '
    'These sizes are easy enough to validate!
    If potentialMatch Then
        
        'Infer dimensions from filesize, then use integer rounding to ensure square dimensions
        m_Width = Int(Sqr(Files.FileLenW(srcFilename) \ 2))
        m_Height = m_Width
        potentialMatch = (Files.FileLenW(srcFilename) = (m_Width * m_Height * 2))
        If potentialMatch Then potentialMatch = (m_Width = 1201) Or (m_Width = 3601)
        
    End If
    
    'Check extension too, as requested.
    If (potentialMatch And requireValidFileExtension) Then
        potentialMatch = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "hgt", True)
    End If
    
    IsFileHGT = potentialMatch
    
End Function

'Validate and load a candidate HGT file
Friend Function LoadHGT_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "LoadHGT_FromFile"
    LoadHGT_FromFile = False
    
    'Validate the file
    If Me.IsFileHGT(srcFile, False) Then
        
        'If validation passed, we already have the (probably correct?) width and height stored in m_width and m_height.
        If (m_Width <= 0) Or (m_Height <= 0) Then
            InternalError FUNC_NAME, "bad dimensions: " & m_Width & "x" & m_Height
            Exit Function
        End If
        
        If HGT_DEBUG_VERBOSE Then
            PDDebug.LogAction "HGT dimensions: " & m_Width & "x" & m_Height
            PDDebug.LogAction "Starting parse..."
        End If
            
        'Open a stream on the target file
        Dim cStream As pdStream
        Set cStream = New pdStream
        
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
            
            'Prep image buffer; we'll dump intensity values straight into it.
            Set dstDIB = New pdDIB
            If dstDIB.CreateBlank(m_Width, m_Height, 32, vbWhite, 255) Then
                
                'Dump the full data-set into a local array, then immediately free the source (we don't need it anymore)
                Dim rawBytes() As Byte
                cStream.ReadBytes rawBytes
                Set cStream = Nothing
                
                'The source data is big-endian ints (elevation, in m) on the range [-32767, 32767]
                ' with -32768 representing data voids.  We want to auto-normalize this, but to do so we first need
                ' to convert the source to little-endian data.
                Dim rawShorts() As Integer, intSA As SafeArray1D
                VBHacks.WrapArrayAroundPtr_Int rawShorts, intSA, VarPtr(rawBytes(0)), m_Width * m_Height * 2
                
                '(While parsing, we're also going to look for max/min values to use in subsequent normalize steps.)
                Dim rawMin As Long, rawMax As Long
                rawMin = 32768
                rawMax = -32769
                
                Dim x As Long, y As Long, xOffset As Long
                For x = 0 To m_Width * m_Height - 1
                    xOffset = x * 2
                    y = rawBytes(xOffset)
                    rawBytes(xOffset) = rawBytes(xOffset + 1)
                    rawBytes(xOffset + 1) = y
                    If (rawShorts(x) < rawMin) Then rawMin = rawShorts(x)
                    If (rawShorts(x) > rawMax) Then rawMax = rawShorts(x)
                Next x
                
                'Min/max values can now be used for normalization, if desired.
                
                'For now, however, ground is automatically colored on the range [0, 5,000m].
                ' (Again, in the future it would be great to give the user control over this.)
                Const GROUND_MAX As Long = 5000, GROUND_MIN As Long = 0
                
                Dim useNormalization As Boolean
                useNormalization = True
                
                Dim groundMin As Single, groundMax As Single, groundNorm As Single
                
                '(We use a 1024 entry gradient LUT for painting; this provides good coverage without overwhelming
                ' a normal-ish cache size.)
                If useNormalization Then
                    groundMin = rawMin
                    groundMax = rawMax
                    groundNorm = 1024! / CSng(rawMax - rawMin)
                Else
                    groundMin = GROUND_MIN
                    groundMax = GROUND_MAX
                    groundNorm = 1024! / CSng(groundMax - groundMin)
                End If
                
                'For fun, I've played with the idea of auto-colorizing based on a standard hypsometric palette:
                ' https://en.wikipedia.org/wiki/Hypsometric_tints
                '
                '...but as that requires user input during loading, I'm sticking with grayscale for now.
                Dim useColor As Boolean
                useColor = False
                
                Dim groundColors() As Long, maxColor As Long, seaColor As Long
                If useColor Then
                    GetGroundGradient_Color groundColors, maxColor, 1024
                    seaColor = Colors.GetRGBALongFromHex("#00a4f0ff")
                Else
                    GetGroundGradient_Gray groundColors, maxColor, 1024
                    seaColor = Colors.GetRGBALongFromHex("#000000ff")
                End If
                
                'Wrap a pixel array around the destination image
                Dim dstPixels() As Long, tmpSA As SafeArray2D
                dstDIB.WrapLongArrayAroundDIB dstPixels, tmpSA
                
                Dim tmpIntensity As Long, curColor As Long
                
                '-32768 means "invalid data"; if encountered, we'll try to reuse the previous value, if any.
                Const HGT_DATA_VOID As Long = -32768
                Dim prevIntensity As Long
                prevIntensity = 0
                
                For y = 0 To m_Height - 1
                For x = 0 To m_Width - 1
                    
                    'Pull the next big-endian value from file and replace invalid entries with previous values
                    tmpIntensity = rawShorts(y * m_Width + x)
                    If (tmpIntensity = HGT_DATA_VOID) Then tmpIntensity = prevIntensity
                    prevIntensity = tmpIntensity
                    
                    'Check max/min values
                    If (tmpIntensity <= groundMin) Then
                        curColor = seaColor
                    ElseIf (tmpIntensity >= groundMax) Then
                        curColor = maxColor
                    Else
                        curColor = groundColors(Int((tmpIntensity - groundMin) * groundNorm + 0.5))
                    End If
                        
                    dstPixels(x, y) = curColor
                    
                Next x
                Next y
                
                VBHacks.UnwrapArrayFromPtr_Int rawShorts
                dstDIB.UnwrapLongArrayFromDIB dstPixels
                LoadHGT_FromFile = True
                
                'The returned data is always premultiplied
                If LoadHGT_FromFile Then dstDIB.SetInitialAlphaPremultiplicationState True
                
                'Regardless of outcome, free the underlying stream
                Set cStream = Nothing
                
            Else
                InternalError FUNC_NAME, "out of memory"
                Set m_Stream = Nothing
                Exit Function
            End If
            
        Else
            InternalError FUNC_NAME, "bad stream"
            Exit Function
        End If
        
    Else
        'File is not HGT; silently ignore it
    End If
    
End Function

'Colors derived from https://www.esri.com/arcgis-blog/products/arcgis-desktop/mapping/creating-a-legend-for-hypsometrically-tinted-shaded-relief/
Private Sub GetGroundGradient_Color(ByRef dstColors() As Long, ByRef maxColor As Long, Optional ByVal gradRange As Long = 1024)
    
    Dim groundSamples() As GradientPoint
    ReDim groundSamples(0 To 9) As GradientPoint
    
    'Set opacity of all points to 100%
    Dim i As Long
    For i = LBound(groundSamples) To UBound(groundSamples)
        groundSamples(i).PointOpacity = 100!
    Next i
    
    'Position is fixed according to elevation, with 5k meters as a fixed upper limit
    groundSamples(0).PointPosition = 0!
    groundSamples(1).PointPosition = 0.2!
    groundSamples(2).PointPosition = 0.3!
    groundSamples(3).PointPosition = 0.4!
    groundSamples(4).PointPosition = 0.5!
    groundSamples(5).PointPosition = 0.6!
    groundSamples(6).PointPosition = 0.7!
    groundSamples(7).PointPosition = 0.8!
    groundSamples(8).PointPosition = 0.9!
    groundSamples(9).PointPosition = 1!
    
    groundSamples(0).PointRGB = RGB(210, 255, 190)
    groundSamples(1).PointRGB = RGB(160, 215, 195)
    groundSamples(2).PointRGB = RGB(135, 205, 100)
    groundSamples(3).PointRGB = RGB(180, 215, 160)
    groundSamples(4).PointRGB = RGB(215, 215, 160)
    groundSamples(5).PointRGB = RGB(245, 245, 120)
    groundSamples(6).PointRGB = RGB(255, 235, 175)
    groundSamples(7).PointRGB = RGB(255, 210, 125)
    groundSamples(8).PointRGB = RGB(215, 175, 160)
    groundSamples(9).PointRGB = RGB(255, 250, 255)
    
    maxColor = groundSamples(9).PointRGB
    
    Dim testGradient As pd2DGradient
    Set testGradient = New pd2DGradient
    testGradient.CreateGradientFromPointCollection 10, groundSamples
    testGradient.GetLookupTable dstColors, gradRange
    
End Sub

Private Sub GetGroundGradient_Gray(ByRef dstColors() As Long, ByRef maxColor As Long, Optional ByVal gradRange As Long = 1024)
    
    Dim testGradient As pd2DGradient
    Set testGradient = New pd2DGradient
    testGradient.CreateTwoPointGradient RGB(0, 0, 0), RGB(255, 255, 255)
    testGradient.GetLookupTable dstColors, gradRange
    
    maxColor = RGB(255, 255, 255)
    
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdHGT." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdHGT." & funcName & "() reported an error: " & errDescription
    End If
End Sub
