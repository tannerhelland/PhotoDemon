VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdLUT3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon 3D LUT I/O class
'Copyright 2020-2022 by Tanner Helland
'Created: 23/October/20
'Last updated: 23/February/22
'Last update: fix parsing issues on weirdly encoded 3rd-party .look files
'
'From the .CUBE spec:
'
'"Look-up tables (LUTs) are often used when converting an image from one color representation
' to another - for example, when converting between log and gamma encodings, changing color space,
' applying color corrections, changing dynamic range, gamut mapping, exporting to an output device,
' or previewing how an image will be reproduced.
'
'"A look-up table is a sampled representation of a mathematical function, with the samples stored
' in table form.  The table form provides a simple and universal representation of arbitrarily
' complex (or proprietary) functions, and doing a table look-up can be faster than calculating the
' the original function.
'
'"Each sample in the table holds the output value(s) for a specific input coordinate. When an
' output value is requested for an input coordinate that falls between stored coordinates, the
' value is obtained by interpolating between adjacent samples."
'
'3D LUTs are an essential aspect of pro workflows, and I'm excited to finally bring support to
' PhotoDemon.  This was a messy project since (obviously) no VB dev has ever tackled this before,
' and some technical aspects - like tetrahedral interpolation - are required by Adobe's spec,
' but they don't actually disclose *how* to do the damn thing.  Thank goodness NVidia has shared a
' Technicolor (yes, *that* Technicolor) presentation on their website that hints at the required math.
'
'Many documents were consulted to build this class.  Critical links include these ones
' (intact as of October 2020):
'
' - Version 1.0 of Adobe's .CUBE spec:
'    https://wwwimages2.adobe.com/content/dam/acom/en/products/speedgrade/cc/pdfs/cube-lut-specification-1.0.pdf
' - Technicolor presentation (via NVidia) on tetrahedral interpolation:
'    https://www.nvidia.com/content/GTC/posters/2010/V01-Real-Time-Color-Space-Conversion-for-High-Resolution-Video.pdf
' - 2009 version of Autodesk's 3DL spec:
'    https://download.autodesk.com/us/systemdocs/pdf/lustre_color_management_user_guide.pdf
'
'This class has been heavily optimized for speed, which means it has many dependencies on other
' PD classes.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'When serializing the LUT array (which is just a large array of floats) to/from a byte stream,
' we can shuffle the floats for greatly improved compression performance.  (Shuffling takes
' the first byte of every float and places them together, then all second bytes, etc; this is
' much more likely to produce useful compression sequences as opposed to normal interlaced data.)
Private Const DEBUG_SHUFFLE_FLOATS As Boolean = True

'Progress reports are returned on the scale (0, 1).  When raised, you can request that the operation
' cancel itself; any in-progress work will be erased.
Event ProgressUpdate(ByVal progressValue As Single, ByRef cancelOperation As Boolean)

'Magic ID when serializing to a byte stream
Private Const PD_LUT_MAGIC_NUMBER As String = "PDLUT3DS"

'Only 3D LUTs are currently supported, at single-precision.
Private Type FloatColor
    rFloat As Single
    gFloat As Single
    bFloat As Single
End Type

Private m_lut3D() As FloatColor

'Path of the currently loaded LUT, if any; null means a LUT has not been loaded yet
Private m_lutPath As String

'Some LUTs (.look format) may reference a sub-file for actual LUT data.  In PD, successfully loaded LUTs
' are typically copied into the local LUT folder to maintain portability.  We need to copy the subfile as
' well, or the LUT won't work.  (We could theoretically work around this by converting the LUT to hex,
' then copying that data into the parent LUT file... but wow, that's a mess.)
Private m_subLutPath As String

'LUT size is uniform for all dimensions, and must be specified by the file (e.g. there is no default value)
' (loop size is just m_LutSize-1, but it's faster to precalculate it since it's accessed inside many loops).
Private m_LutSize As Long, m_LutLoopSize As Long

'Domain min and max values can vary; these are typically [0, 1] but the spec allows any arbitrary value.
' (Proliferation of HDR compatibility may mean it's more likely to see values outside the normal range
' in the future.  This class handles such tables just fine.)
Private m_MinR As Single, m_MinG As Single, m_MinB As Single
Private m_MaxR As Single, m_MaxG As Single, m_MaxB As Single
Private m_RangeR As Single, m_RangeG As Single, m_RangeB As Single

'Some luts are loaded across multiple stages.  State-tracking is important to prevent errors.
Private m_LoadSuccessful As Boolean

'Apply the loaded lUT to an arbitrary DIB
Friend Sub ApplyLUTToDIB(ByRef srcDIB As pdDIB, Optional ByVal raiseProgressEvents As Boolean = False)
    
    'This filter is interruptible
    Dim cancelOp As Boolean
    cancelOp = False
    
    'If the caller wants progress reports, calculate some max/min progress values
    Dim progCheck As Long, progMax As Single
    If raiseProgressEvents Then
        progMax = srcDIB.GetDIBHeight
        progCheck = ProgressBars.FindBestProgBarValue(progMax)
        progMax = (1! / progMax)
    End If
    
    'Unpremultiply the source DIB's alpha, as necessary
    Dim srcWasPremultiplied As Boolean
    If (srcDIB.GetDIBColorDepth = 32) And srcDIB.GetAlphaPremultiplication() Then
        srcWasPremultiplied = True
        srcDIB.SetAlphaPremultiplication False
    Else
        srcWasPremultiplied = False
    End If
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0 * 4
    initY = 0
    finalX = (srcDIB.GetDIBWidth - 1) * 4
    finalY = srcDIB.GetDIBHeight - 1
    
    'Color values are used, but I plan on doing some testing to see if direct LUT assignment is faster.
    Dim r As Long, g As Long, b As Long, newR As Long, newG As Long, newB As Long
    
    'Now we can loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Apply the active LUT
        Me.GetInterpolatedRGB_Int r, g, b, newR, newG, newB
        
        'Assign the look-up table values
        imageData(x) = newB
        imageData(x + 1) = newG
        imageData(x + 2) = newR
        
    Next x
    
        'UI updates
        If raiseProgressEvents Then
            If ((y And progCheck) = 0) Then
                RaiseEvent ProgressUpdate((CDbl(y) * progMax), cancelOp)
                If cancelOp Then GoTo FilterCleanup
            End If
        End If
        
    Next y

FilterCleanup:

    'Safely deallocate imageData()
    srcDIB.UnwrapArrayFromDIB imageData
    
    'Reset alpha premultiplication, as necessary
    If srcWasPremultiplied Then srcDIB.SetAlphaPremultiplication True
    
End Sub

'Display PD's generic 3D LUT load dialog.  All supported LUT filetypes will be available to the user.
' Note that this function returns TRUE if the user selects a file, FALSE if they cancel.  This function
' DOES NOT attempt to load or validate the target file - that's on you.
Public Function DisplayLUTLoadDialog(ByRef srcFilename As String, ByRef dstFilename As String) As Boolean
    
    DisplayLUTLoadDialog = False
    
    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    Dim cdFilter As pdString
    Set cdFilter = New pdString
    cdFilter.Append g_Language.TranslateMessage("All supported LUTs") & "|*.cube;*.look;*.3dl|"
    
    cdFilter.Append "Adobe / IRIDAS (.cube)|*.cube|"
    cdFilter.Append "Adobe SpeedGrade (.look)|*.look|"
    cdFilter.Append "Autodesk Lustre (.3dl)|*.3dl|"
    cdFilter.Append g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Select a LUT file")
    
    'Prep a common dialog interface
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
            
    Dim sFile As String
    sFile = srcFilename
    
    If openDialog.GetOpenFileName(sFile, , True, False, cdFilter.ToString(), 1, UserPrefs.GetLUTPath(False), cdTitle, , GetModalOwner().hWnd) Then
    
        'By design, we don't perform any validation here.  Let the caller validate the file as much (or as little)
        ' as they require.
        DisplayLUTLoadDialog = (LenB(sFile) <> 0)
        
        'The dialog was successful.  Return the path, and save this path for future usage.
        If DisplayLUTLoadDialog Then
            UserPrefs.SetLUTPath sFile
            dstFilename = sFile
        Else
            dstFilename = vbNullString
        End If
        
    End If
    
    'Re-enable user input
    Interface.EnableUserInput
    
End Function

Friend Function GetLUTPath() As String
    GetLUTPath = m_lutPath
End Function

'Only some files (.look format) require a subpath
Friend Function GetLUTSubPath() As String
    GetLUTSubPath = m_subLutPath
End Function

'Load any LUT from file by calling this function.  The LUT format doesn't need to be known in advance;
' we'll sort it out based on file extension.  (A necessary evil since these formats are all just text files.)
Friend Function LoadLUTFromFile(ByRef srcFile As String) As Boolean

    Dim strExtension As String
    strExtension = LCase$(Files.FileGetExtension(srcFile))
    
    If (strExtension = "cube") Then
        LoadLUTFromFile = LoadLUTFromFile_CUBE(srcFile)
    ElseIf (strExtension = "3dl") Then
        LoadLUTFromFile = LoadLUTFromFile_3DL(srcFile)
    ElseIf (strExtension = "look") Then
        LoadLUTFromFile = LoadLUTFromFile_look(srcFile)
    Else
        InternalProblem "LoadLUTFromFile", "unknown file format"
        LoadLUTFromFile = False
    End If
    
    'Do *NOT* overwrite the LUT path if this is a sub-file
    If LoadLUTFromFile Then
        m_lutPath = srcFile
    Else
        m_lutPath = vbNullString
        InternalProblem "LoadLUTFromFile", "load failure: " & srcFile
    End If

End Function

'3DL spec here: https://download.autodesk.com/us/systemdocs/pdf/lustre_color_management_user_guide.pdf
Private Function LoadLUTFromFile_3DL(ByRef srcFile As String) As Boolean

    On Error GoTo LUTParseCrash
    Me.ResetLUT
    
    Dim srcStream As pdStream
    Set srcStream = New pdStream
    If srcStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
    
        '3DL files are human-readable text files.  Load the entire thing into a string array.
        Dim rawFileString As String
        srcStream.SetPosition 0, FILE_BEGIN
        rawFileString = srcStream.ReadString_UnknownEncoding(srcStream.GetStreamSize(), False)
        If (LenB(rawFileString) <> 0) Then
            
            'Assume the rest of the file will parse correctly (if it doesn't, we'll reset this)
            m_LoadSuccessful = True
            
            'The spec is ambiguous about what constitutes a "delimiter".  I've only seen spaces
            ' in the wild, but just in case, check for tabs and convert if found.
            Dim spaceChar As String
            spaceChar = " "
            If (InStr(1, rawFileString, vbTab, vbBinaryCompare) <> 0) Then rawFileString = Replace$(rawFileString, vbTab, spaceChar, 1, -1, vbBinaryCompare)
            
            '3DL files are also a little different from CUBE files because they can use vbCrLf
            ' as a line delimiter (instead of vbLf); check for this and remove all carriage returns,
            ' so we can parse this as similarly to a CUBE file as possible.
            If (InStr(1, rawFileString, vbCr, vbBinaryCompare) <> 0) Then rawFileString = Replace$(rawFileString, vbCr, vbNullString, 1, -1, vbBinaryCompare)
            
            'For perf reasons, we use a pdString object to parse the incoming string into
            ' individual lines.
            Dim fastParse As pdString
            Set fastParse = New pdString
            fastParse.Append rawFileString
            
            'Some complex lines are parsed using a string stack object
            Dim wordsInLine As pdStringStack
            Set wordsInLine = New pdStringStack
            
            'When we hit a value line, all remaining lines in the file must be value lines *ONLY*.
            ' (This is also a useful optimization, as it allows us to bypass other perf-intensive checks.)
            Dim firstLineHit As Boolean, headerHit As Boolean
            firstLineHit = False
            headerHit = False
            
            Dim rCounter As Long, gCounter As Long, bCounter As Long
            
            'Parse each line in turn.
            Dim i As Long, j As Long, k As Long
            Dim curLine As String, curLineLen As Long
            
            Dim linePos As Long, startPos As Long, endPos As Long
            linePos = 1
            Do While fastParse.GetLine(curLine, linePos, vbLf)
                
                'Before processing the current line, calculate the next line's start position
                linePos = linePos + Len(curLine) + 1
                
                '(Crashing on a LUT?  Uncomment to see what we're attempting to parse:
                'Debug.Print "|" & curLine & "|"
                
                'Trim leading and trailing whitespace from all lines before parsing
                curLine = Trim$(curLine)
                curLineLen = Len(curLine)
                
                'The smallest useful line size in a .3dl file is 5 characters:
                '"0 0 0"
                'As such, we can skip any line with less than 5 characters.
                '(We can also skip lines that start with a #; these are comments.)
                If (curLineLen >= 5) Then
                If (AscW(curLine) <> 35) Then
                    
                    'As an optimization, if we're already in the "values" portion of the file,
                    ' skip a bunch of other checks and start parsing values directly.
                    ' (Note: if VB6 supported any notion of "inlining", this would be a separate
                    ' function... but alas we don't live in that timeline.  So instead, I've done
                    ' some ugly copy+paste for perf reasons.)
                    If firstLineHit Then
                        
                        'Look for the first and second occurrences of "space" characters in the
                        ' target string.  Use these to splice the string into thirds, and treat
                        ' each third as a unique floating-point value.
                        startPos = InStr(1, curLine, spaceChar, vbBinaryCompare)
                        m_lut3D(rCounter, gCounter, bCounter).rFloat = Val(Left$(curLine, startPos - 1))
                        endPos = InStr(startPos + 1, curLine, spaceChar, vbBinaryCompare)
                        m_lut3D(rCounter, gCounter, bCounter).gFloat = Val(Mid$(curLine, startPos + 1, endPos - startPos))
                        m_lut3D(rCounter, gCounter, bCounter).bFloat = Val(Right$(curLine, curLineLen - endPos))
                        
                        'Increment counters
                        bCounter = bCounter + 1
                        If (bCounter >= m_LutSize) Then
                            bCounter = 0
                            gCounter = gCounter + 1
                        End If
                        If (gCounter >= m_LutSize) Then
                            gCounter = 0
                            rCounter = rCounter + 1
                            
                            'Some 3DL files add a bunch of useless shit after the final data point,
                            ' so just to be safe, stop parsing after we hit the final mandatory value
                            If (rCounter >= m_LutSize) Then Exit Do
                            
                        End If
                        
                    'We haven't hit a value line, which means we are still inside the "header" region
                    ' of the file.
                    Else
                        
                        'The first non-null, non-comment line in the file is a list of evenly spaced
                        ' "node" coordinates for this particular transform.  The spec is terribly written
                        ' on this point, e.g. here is what it says:
                        
                        '"The first line of a 3D LUT file determines the level of segmentation of the color
                        ' cube, which determines the number of vertices to which you can assign values.
                        ' While each color channel could be segmented differently, a uniform RGB
                        ' segmentation of 17 is supported. The first line of the 3D LUT, then, indicates at
                        ' what 17 intervals input and output values are matched. All RGB values are
                        ' represented in the 10-bit color space. So, the first line of a 3D LUT indicates
                        ' the 17 sampling intervals in the 10-bit space: 0 64 128 192 256 320 384 448 512
                        ' 576 640 704 768 832 896 960 1023."
                        
                        'I have examined a number of "in-the-wild" 3dl files, and none have (so far)
                        ' deviated from this example of 16 sampling intervals.  (Some have off-by-one
                        ' errors which seems unintentional rather than meaningful.)  Rather than assume
                        ' that *all* files use 17 intervals, our parser will actually read the first
                        ' and last values to determine a total range, and it will count the number of
                        ' vertices in case it is *not* 17 - but it will not, however, support
                        ' unevenly-spaced vertices.  (This would be an interpolation nightmare.)
                        If (Not headerHit) Then
                            
                            'Perform a basic sanity check
                            If (Left$(curLine, 2) <> "0 ") Then GoTo BadLineValue
                            
                            'This must be the header.
                            If firstLineHit Then InternalProblem "LoadLUTFromFile", "bad line order"
                            
                            'Parse the list of values into a list of discrete "words".
                            Set wordsInLine = Strings.GetListOfWordsFromString(curLine)
                            
                            'We really only care about the number of discrete values, because that's
                            ' what determines the size of our lookup table.
                            m_LutSize = wordsInLine.GetNumOfStrings()
                            m_LutLoopSize = m_LutSize - 1
                            
                            'If this format made any sense, we would determine min/max values by
                            ' retrieving them from the input table, e.g.
                            'm_MinR = Val(wordsInLine.GetString(0))
                            'm_MaxR = Val(wordsInLine.GetString(m_LutLoopSize))
                            '...but from everything I can see online, these files instead seem to always
                            ' map to 12-bit values on the range [0, 4095], so we do that too.  If people
                            ' can show me LUT files that use a different remap range, I can add heuristics
                            ' to solve this more elegantly, but for now, I don't want to spend too much time
                            ' messing with it.
                            m_MinR = 0
                            m_MinG = m_MinR
                            m_MinB = m_MinR
                            
                            m_MaxR = 4095
                            m_MaxG = m_MaxR
                            m_MaxB = m_MaxR
                            
                            'Note that the header has been hit, which means we shouldn't look for it again
                            headerHit = True
                        
                        'This must be the first value line!
                        Else
                        
                            'Note that we've hit a value line.  All lines beyond this point in the file
                            ' must also be value lines.
                            firstLineHit = True
                            
                            'Prepare the LUT table and reset all counters
                            ReDim m_lut3D(0 To m_LutLoopSize, 0 To m_LutLoopSize, 0 To m_LutLoopSize) As FloatColor
                            rCounter = 0
                            gCounter = 0
                            bCounter = 0
                            
                            'Retrieve the first RGB triplet
                            startPos = InStr(1, curLine, spaceChar, vbBinaryCompare)
                            m_lut3D(rCounter, gCounter, bCounter).rFloat = Val(Left$(curLine, startPos - 1))
                            endPos = InStr(startPos + 1, curLine, spaceChar, vbBinaryCompare)
                            m_lut3D(rCounter, gCounter, bCounter).gFloat = Val(Mid$(curLine, startPos + 1, endPos - startPos))
                            m_lut3D(rCounter, gCounter, bCounter).bFloat = Val(Right$(curLine, Len(curLine) - endPos))
                            
                            'Increment the blue counter so we're ready for the next point [0, 0, 1]
                            bCounter = bCounter + 1
                            
                        End If
                        
                    End If
                    
                '/end "line is at least five chars long" and "line is not a comment"
                End If
                End If
                
            'Continue parsing remaining lines
BadLineValue:
            Loop
                            
            If m_LoadSuccessful Then
                
                'Calculate min/max ranges
                m_RangeR = m_MaxR - m_MinR
                m_RangeG = m_MaxG - m_MinG
                m_RangeB = m_MaxB - m_MinB
                
                'If the minimum isn't 0, we need to subtract it from all values to re-base them
                If (m_MinR <> 0!) Or (m_MinG <> 0!) Or (m_MinB <> 0!) Then
                    
                    For i = 0 To m_LutLoopSize
                        For j = 0 To m_LutLoopSize
                            For k = 0 To m_LutLoopSize
                                m_lut3D(i, j, k).rFloat = m_lut3D(i, j, k).rFloat - m_MinR
                                m_lut3D(i, j, k).gFloat = m_lut3D(i, j, k).gFloat - m_MinG
                                m_lut3D(i, j, k).bFloat = m_lut3D(i, j, k).bFloat - m_MinB
                            Next k
                        Next j
                    Next i
                    
                End If
                
                'If the range isn't 1!, we need to divide all values so that they're on the range [0, 1]
                If (m_RangeR <> 1!) Or (m_RangeG <> 1!) Or (m_RangeB <> 1!) Then
                    
                    Dim invR As Single, invG As Single, invB As Single
                    If (m_RangeR <> 0!) Then invR = 1! / m_RangeR
                    If (m_RangeG <> 0!) Then invG = 1! / m_RangeG
                    If (m_RangeB <> 0!) Then invB = 1! / m_RangeB
                    
                    For i = 0 To m_LutLoopSize
                        For j = 0 To m_LutLoopSize
                            For k = 0 To m_LutLoopSize
                                m_lut3D(i, j, k).rFloat = m_lut3D(i, j, k).rFloat * invR
                                m_lut3D(i, j, k).gFloat = m_lut3D(i, j, k).gFloat * invG
                                m_lut3D(i, j, k).bFloat = m_lut3D(i, j, k).bFloat * invB
                            Next k
                        Next j
                    Next i
                    
                End If
                
            End If
            
        Else
            m_LoadSuccessful = False
            GoTo BadLUTFile
        End If
        
    Else
        m_LoadSuccessful = False
        GoTo BadLUTFile
    End If
    
BadLUTFile:
    LoadLUTFromFile_3DL = m_LoadSuccessful
    Exit Function
    
LUTParseCrash:
    InternalProblem "LoadLUTFromFile3DL", "VB Error # " & Err.Number & ": " & Err.Description
    LoadLUTFromFile_3DL = False
End Function

'Adobe CUBE format
Private Function LoadLUTFromFile_CUBE(ByRef srcFile As String) As Boolean

    On Error GoTo LUTParseCrash
    Me.ResetLUT
    
    Dim srcStream As pdStream
    Set srcStream = New pdStream
    If srcStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
    
        '.CUBE files are human-readable text files.  Load the entire thing into a string array.
        Dim rawFileString As String
        srcStream.SetPosition 0, FILE_BEGIN
        rawFileString = srcStream.ReadString_UnknownEncoding(srcStream.GetStreamSize(), False)
        If (LenB(rawFileString) <> 0) And (Strings.StrStrBM(rawFileString, "LUT_3D_SIZE") <> 0) Then
            
            'Assume the rest of the file will parse correctly (if it doesn't, we'll reset this)
            m_LoadSuccessful = True
            
            'The spec allows tab characters or space characters to be used as delimiters.
            ' We only want to support space characters, as it improves parse performance.
            Dim spaceChar As String
            spaceChar = " "
            If (InStr(1, rawFileString, vbTab, vbBinaryCompare) <> 0) Then rawFileString = Replace$(rawFileString, vbTab, spaceChar, 1, -1, vbBinaryCompare)
            
            'For perf reasons, we use a pdString object to parse the incoming string into
            ' individual lines.
            Dim fastParse As pdString
            Set fastParse = New pdString
            fastParse.Append rawFileString
            
            'Some complex lines are parsed using a string stack object
            Dim wordsInLine As pdStringStack
            Set wordsInLine = New pdStringStack
            
            'When we hit a value line, all remaining lines in the file must be value lines *ONLY*.
            ' (This is also a useful optimization, as it allows us to bypass other perf-intensive checks.)
            Dim firstLineHit As Boolean
            firstLineHit = False
            
            Dim rCounter As Long, gCounter As Long, bCounter As Long
            
            'Parse each line in turn.
            Dim i As Long, j As Long, k As Long
            Dim curLine As String, curLineLen As Long
            
            Dim linePos As Long, startPos As Long, endPos As Long
            linePos = 1
            Do While fastParse.GetLine(curLine, linePos, vbLf)
                
                'Before processing the current line, calculate the next line's start position
                linePos = linePos + Len(curLine) + 1
                
                'Trim leading and trailing whitespace from all lines before parsing
                curLine = Trim$(curLine)
                curLineLen = Len(curLine)
                
                'The smallest useful line size in a .CUBE file is 5 characters:
                '"0 0 0"
                'As such, we can skip any line with less than 5 characters.
                '(We can also skip lines that start with a #; these are comments.)
                If (curLineLen >= 5) Then
                If (AscW(curLine) <> 35) Then
                    
                    'As an optimization, if we're already in the "values" portion of the file,
                    ' skip a bunch of other checks and start parsing values directly.
                    ' (Note: if VB6 supported any notion of "inlining", this would be a separate
                    ' function... but alas we don't live in that timeline.  So instead, I've done
                    ' some ugly copy+paste for perf reasons.)
                    If firstLineHit Then
                        
                        'Look for the first and second occurrences of "space" characters in the
                        ' target string.  Use these to splice the string into thirds, and treat
                        ' each third as a unique floating-point value.
                        startPos = InStr(1, curLine, spaceChar, vbBinaryCompare)
                        m_lut3D(rCounter, gCounter, bCounter).rFloat = Val(Left$(curLine, startPos - 1))
                        endPos = InStr(startPos + 1, curLine, spaceChar, vbBinaryCompare)
                        m_lut3D(rCounter, gCounter, bCounter).gFloat = Val(Mid$(curLine, startPos + 1, endPos - startPos))
                        m_lut3D(rCounter, gCounter, bCounter).bFloat = Val(Right$(curLine, curLineLen - endPos))
                        
                        'Increment counters
                        rCounter = rCounter + 1
                        If (rCounter >= m_LutSize) Then
                            rCounter = 0
                            gCounter = gCounter + 1
                        End If
                        If (gCounter >= m_LutSize) Then
                            gCounter = 0
                            bCounter = bCounter + 1
                            If (bCounter >= m_LutSize) Then Exit Do
                        End If
                    
                    'We haven't hit a value line, which means we are still inside the "header" region
                    ' of the file.
                    Else
                        
                        'Lines with certain prefixes have special meaning
                        If Strings.LeftMatches(curLine, "LUT_3D_SIZE", True) Then
                            If firstLineHit Then InternalProblem "LoadLUTFromFile", "bad line order"
                            m_LutSize = Strings.RightByChar(curLine, spaceChar)
                            m_LutLoopSize = m_LutSize - 1
                            
                        'Minimum values, typically 0.0
                        ElseIf Strings.LeftMatches(curLine, "DOMAIN_MIN", True) Then
                            If firstLineHit Then InternalProblem "LoadLUTFromFile", "bad line order"
                            
                            'Retrieve all three values
                            Set wordsInLine = Strings.GetListOfWordsFromString(curLine)
                            m_MinR = wordsInLine.GetString(1)
                            m_MinG = wordsInLine.GetString(2)
                            m_MinB = wordsInLine.GetString(3)
                            
                        ElseIf Strings.LeftMatches(curLine, "DOMAIN_MAX", True) Then
                            If firstLineHit Then InternalProblem "LoadLUTFromFile", "bad line order"
                            
                            Set wordsInLine = Strings.GetListOfWordsFromString(curLine)
                            m_MaxR = wordsInLine.GetString(1)
                            m_MaxG = wordsInLine.GetString(2)
                            m_MaxB = wordsInLine.GetString(3)
                            
                        ElseIf (Left$(curLine, 5) = "TITLE") Then
                            'Retrieving the title (if any) is TODO
                            
                        'Assume any other lines are valid table inputs
                        Else
                            
                            'Note that we've hit a value line.  All lines beyond this point in the file
                            ' must also be value lines.
                            firstLineHit = True
                            
                            'Prepare the LUT table and reset all counters
                            ReDim m_lut3D(0 To m_LutLoopSize, 0 To m_LutLoopSize, 0 To m_LutLoopSize) As FloatColor
                            rCounter = 0
                            gCounter = 0
                            bCounter = 0
                            
                            'Retrieve the first RGB triplet
                            startPos = InStr(1, curLine, spaceChar, vbBinaryCompare)
                            m_lut3D(rCounter, gCounter, bCounter).rFloat = Val(Left$(curLine, startPos - 1))
                            endPos = InStr(startPos + 1, curLine, spaceChar, vbBinaryCompare)
                            m_lut3D(rCounter, gCounter, bCounter).gFloat = Val(Mid$(curLine, startPos + 1, endPos - startPos))
                            m_lut3D(rCounter, gCounter, bCounter).bFloat = Val(Right$(curLine, Len(curLine) - endPos))
                            
                            'Increment the red counter so we're ready for the next point [1, 0, 0]
                            rCounter = rCounter + 1
                            
                        End If
                        
                    End If
                    
                '/end "line is at least five chars long" and "line is not a comment"
                End If
                End If
                
            'Continue parsing remaining lines
BadLineValue:
            Loop
                            
            If m_LoadSuccessful Then
                
                'Calculate min/max ranges
                m_RangeR = m_MaxR - m_MinR
                m_RangeG = m_MaxG - m_MinG
                m_RangeB = m_MaxB - m_MinB
                
                'If the minimum isn't 0, we need to subtract it from all values to re-base them
                If (m_MinR <> 0!) Or (m_MinG <> 0!) Or (m_MinB <> 0!) Then
                    
                    For i = 0 To m_LutLoopSize
                        For j = 0 To m_LutLoopSize
                            For k = 0 To m_LutLoopSize
                                m_lut3D(i, j, k).rFloat = m_lut3D(i, j, k).rFloat - m_MinR
                                m_lut3D(i, j, k).gFloat = m_lut3D(i, j, k).gFloat - m_MinG
                                m_lut3D(i, j, k).bFloat = m_lut3D(i, j, k).bFloat - m_MinB
                            Next k
                        Next j
                    Next i
                    
                End If
                
                'If the range isn't 1!, we need to divide all values so that they're on the range [0, 1]
                If (m_RangeR <> 1!) Or (m_RangeG <> 1!) Or (m_RangeB <> 1!) Then
                    
                    Dim invR As Single, invG As Single, invB As Single
                    If (m_RangeR <> 0!) Then invR = 1! / m_RangeR
                    If (m_RangeG <> 0!) Then invG = 1! / m_RangeG
                    If (m_RangeB <> 0!) Then invB = 1! / m_RangeB
                    
                    For i = 0 To m_LutLoopSize
                        For j = 0 To m_LutLoopSize
                            For k = 0 To m_LutLoopSize
                                m_lut3D(i, j, k).rFloat = m_lut3D(i, j, k).rFloat * invR
                                m_lut3D(i, j, k).gFloat = m_lut3D(i, j, k).gFloat * invG
                                m_lut3D(i, j, k).bFloat = m_lut3D(i, j, k).bFloat * invB
                            Next k
                        Next j
                    Next i
                    
                End If
                
            End If
            
        Else
            m_LoadSuccessful = False
            GoTo BadLUTFile
        End If
        
    Else
        m_LoadSuccessful = False
        GoTo BadLUTFile
    End If
    
BadLUTFile:
    LoadLUTFromFile_CUBE = m_LoadSuccessful
    Exit Function
    
LUTParseCrash:
    InternalProblem "LoadLUTFromFileCube", "VB Error # " & Err.Number & ": " & Err.Description
    LoadLUTFromFile_CUBE = False
End Function

'Adobe (?) LOOK format
Private Function LoadLUTFromFile_look(ByRef srcFile As String) As Boolean

    On Error GoTo LUTParseCrash
    Me.ResetLUT
    
    Dim srcStream As pdStream
    Set srcStream = New pdStream
    If srcStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
    
        '.look files are just XML.  Load the entire thing into a string.
        Dim rawFileString As String
        srcStream.SetPosition 0, FILE_BEGIN
        rawFileString = srcStream.ReadString_UnknownEncoding(srcStream.GetStreamSize(), False)
        If (LenB(rawFileString) <> 0) And (Strings.StrStrBM(rawFileString, "<LUT>") <> 0) Then
            
            'Let's cheat and use PD's lightweight "fake" XML parser to retrieve just the LUT data
            Dim cXML As pdSerialize
            Set cXML = New pdSerialize
            cXML.SetParamString rawFileString
            
            'Now, a bit of hackery.  I encountered an online LUT collection that included a bunch of
            ' null "LUT" tags inside the preliminary shader segment of the .look file.  I don't know how
            ' they managed to do this but it's probably the result of a garbage 3rd-party LUT generator.
            '
            'Rather than do a full-blown XML parse (which would require more work than I care to implement),
            ' look for the *last* LUT tag in the file.  This will be the correct one.
            Dim lutText As String, tmpPosition As Long
            tmpPosition = InStrRev(rawFileString, "<LUT>", -1, vbTextCompare)
            If (tmpPosition <= 0) Then tmpPosition = -1
            
            If cXML.GetStringEx("LUT", lutText, True, tmpPosition) Then
                
                'We need two values from the LUT section.
                cXML.SetParamString lutText
                
                'First, we need to know the size of the LUT
                Dim sizeString As String
                sizeString = cXML.GetString("size", vbNullString, True)
                
                'Some special .look LUT files (described below) won't have a size parameter;
                ' we'll deal with these in a moment.
                Dim noSize As Boolean
                noSize = (LenB(sizeString) <= 0)
                
                If (Not noSize) Then
                    
                    'Size is always an integer delimited by quotation marks
                    sizeString = Replace(sizeString, """", vbNullString)
                    
                    Dim sizeInt As Long
                    sizeInt = Val(sizeString)
                    If (sizeInt < 0) Then Exit Function
                    
                    m_LutSize = sizeInt
                    m_LutLoopSize = m_LutSize - 1
                    
                End If
                    
                'Next, we need the LUT data itself
                Dim lutString As String
                lutString = cXML.GetString("data", vbNullString, True)
                
                'Remove all useless formatting chars
                lutString = Replace(lutString, """", vbNullString)
                lutString = Replace(lutString, vbTab, vbNullString)
                lutString = Replace(lutString, vbCr, vbNullString)
                lutString = Replace(lutString, vbLf, vbNullString)
                
                'Now we have to address a rare but valid situation.  Some .look files don't
                ' actually embed full LUT data.  Instead, they simply point at another LUT
                ' file in the same folder.  We want to try and detect this case, and fall back
                ' to a separate load call as necessary.
                
                'The easiest way to catch this case is to simply calculate the size of the
                ' LUT table we expect to see.  If the returned string is smaller than this,
                ' check for a filename.
                Dim minRequiredSize As Long
                minRequiredSize = m_LutSize * m_LutSize * m_LutSize * 3 * 8  '3 color components * 8 hex chars required to describe 4 bytes worth of data per component
                
                If (noSize Or (Len(lutString) < minRequiredSize - 1)) Then
                
                    'Uh oh - we don't have enough data to parse hex data.  See if this is a filename instead.
                    Dim potentialFile As String
                    potentialFile = Files.PathAddBackslash(Files.FileGetPath(srcFile)) & lutString
                    
                    Dim subFileRequired As Boolean
                    subFileRequired = Files.FileExists(potentialFile)
                    
                    'Some files may not list the file in the <Data> segment, but may instead place it directly
                    ' in the LUT tag.
                    If (Not subFileRequired) Then
                        potentialFile = Files.PathAddBackslash(Files.FileGetPath(srcFile)) & Replace(lutText, """", vbNullString)
                        subFileRequired = Files.FileExists(potentialFile)
                    End If
                    
                    'If this file simply references another external LUT file (sigh), we want to load that LUT file,
                    ' then copy its contents into *this* LUT class.  This allows us to support nested LUT files
                    ' without complex detection criteria.
                    If subFileRequired Then
                        
                        'Cache a copy of the sub-LUT file's location; we need to restore this because the next
                        ' step we perform (serialization) will reset many of this class's module-level values.
                        Dim cChild As pdLUT3D
                        Set cChild = New pdLUT3D
                        m_LoadSuccessful = cChild.LoadLUTFromFile(potentialFile)
                        If m_LoadSuccessful Then
                        
                            'Overwrite this LUT's data with the data from its child
                            Dim lutBytes() As Byte, lutSize As Long
                            cChild.Serialize_ToBytes lutBytes, lutSize, cf_None
                            Me.Serialize_FromBytes lutBytes, lutSize, lutSize, cf_None
                            
                            'Restore the sub-LUT path (as it will have been erased by the serialize command, above)
                            m_subLutPath = potentialFile
                            
                            'Note that m_LUTPath will also be null at this point - that's okay, because the parent
                            ' Load_XYZ function sets that path before exiting when a file loads correctly.
                            
                        Else
                            m_subLutPath = vbNullString
                        End If
                        
                    End If
                
                Else
                    
                    lutString = Replace(lutString, " ", vbNullString)
                    
                    'Attempt to convert the length list of hexes to a standard byte array
                    'Dim lutBytes() As Byte
                    If Strings.BytesFromHex(lutBytes, lutString) Then
                    
                        'If we're still here, the result is likely useable!
                        m_LoadSuccessful = True
                        
                        'Initialize all key LUT values
                        m_MinR = 0!
                        m_MinG = m_MinR
                        m_MinB = m_MinR
                        m_MaxR = 1!
                        m_MaxG = m_MaxR
                        m_MaxB = m_MaxR
                        m_RangeR = m_MaxR - m_MinR
                        m_RangeG = m_MaxG - m_MinG
                        m_RangeB = m_MaxB - m_MinB
                        
                        'Initialize the LUT array
                        ReDim m_lut3D(0 To m_LutLoopSize, 0 To m_LutLoopSize, 0 To m_LutLoopSize) As FloatColor
                        
                        'Copy the bytes as-is into the float array
                        CopyMemoryStrict VarPtr(m_lut3D(0, 0, 0)), VarPtr(lutBytes(0)), m_LutSize * m_LutSize * m_LutSize * 4 * 3
                        
                    End If
                    
                End If
                    
            '/end load LUT section
            End If
            
        Else
            m_LoadSuccessful = False
            GoTo BadLUTFile
        End If
        
    Else
        m_LoadSuccessful = False
        GoTo BadLUTFile
    End If
    
BadLUTFile:
    LoadLUTFromFile_look = m_LoadSuccessful
    Exit Function
    
LUTParseCrash:
    InternalProblem "LoadLUTFromFile_look", "VB Error # " & Err.Number & ": " & Err.Description
    LoadLUTFromFile_look = False
End Function

Friend Sub GetInterpolatedRGB_Int(ByVal srcR As Long, ByVal srcG As Long, ByVal srcB As Long, ByRef dstR As Long, ByRef dstG As Long, ByRef dstB As Long)

    'Use the source RGB values as indices into the master lookup table
    Const ONE_DIV_255 As Single = 1! / 255!
    Dim rIndex As Single, gIndex As Single, bIndex As Single
    rIndex = (srcR * m_LutLoopSize) * ONE_DIV_255
    gIndex = (srcG * m_LutLoopSize) * ONE_DIV_255
    bIndex = (srcB * m_LutLoopSize) * ONE_DIV_255
    
    'Calculate the fractional component of each index; we need this to interpolate.
    ' (While here, we also calculate a "+1" index; this lets us skip bounds-checking on the inner
    ' interpolation loop.)
    Dim nR As Long, nG As Long, nB As Long
    nR = Int(rIndex)
    nG = Int(gIndex)
    nB = Int(bIndex)
    
    Dim nRP As Long, nGP As Long, nBP As Long
    nRP = nR + 1
    If (nRP >= m_LutSize) Then nRP = m_LutLoopSize
    nGP = nG + 1
    If (nGP >= m_LutSize) Then nGP = m_LutLoopSize
    nBP = nB + 1
    If (nBP >= m_LutSize) Then nBP = m_LutLoopSize
    
    Dim fR As Single, fG As Single, fB As Single
    fR = rIndex - nR
    fG = gIndex - nG
    fB = bIndex - nB
    
    Dim dstRFloat As Single, dstGFloat As Single, dstBFloat As Single
    
    'Per nvidia:
    ' fG>=fB>=fR
    If (fG >= fB) And (fB >= fR) Then
        '(1-fG)·[nR;nG;nB] + (fG-fB)·[nR;nG+1;nB]  +(fB-fR)·[nR;nG+1;nB+1] + (fR)·[nR+1;nG+1;nB+1]
        dstRFloat = (1! - fG) * m_lut3D(nR, nG, nB).rFloat + (fG - fB) * m_lut3D(nR, nGP, nB).rFloat + (fB - fR) * m_lut3D(nR, nGP, nBP).rFloat + fR * m_lut3D(nRP, nGP, nBP).rFloat
        dstGFloat = (1! - fG) * m_lut3D(nR, nG, nB).gFloat + (fG - fB) * m_lut3D(nR, nGP, nB).gFloat + (fB - fR) * m_lut3D(nR, nGP, nBP).gFloat + fR * m_lut3D(nRP, nGP, nBP).gFloat
        dstBFloat = (1! - fG) * m_lut3D(nR, nG, nB).bFloat + (fG - fB) * m_lut3D(nR, nGP, nB).bFloat + (fB - fR) * m_lut3D(nR, nGP, nBP).bFloat + fR * m_lut3D(nRP, nGP, nBP).bFloat
        
    ' fB>fR>fG
    ElseIf (fB > fR) And (fR > fG) Then
        '(1-fB)·[nR;nG;nB] + (fB-fR)·[nR;nG;nB+1] +(fR-fG)·[nR+1;nG;nB+1] + (fG)·[nR+1;nG+1;nB+1]
        dstRFloat = (1! - fB) * m_lut3D(nR, nG, nB).rFloat + (fB - fR) * m_lut3D(nR, nG, nBP).rFloat + (fR - fG) * m_lut3D(nRP, nG, nBP).rFloat + fG * m_lut3D(nRP, nGP, nBP).rFloat
        dstGFloat = (1! - fB) * m_lut3D(nR, nG, nB).gFloat + (fB - fR) * m_lut3D(nR, nG, nBP).gFloat + (fR - fG) * m_lut3D(nRP, nG, nBP).gFloat + fG * m_lut3D(nRP, nGP, nBP).gFloat
        dstBFloat = (1! - fB) * m_lut3D(nR, nG, nB).bFloat + (fB - fR) * m_lut3D(nR, nG, nBP).bFloat + (fR - fG) * m_lut3D(nRP, nG, nBP).bFloat + fG * m_lut3D(nRP, nGP, nBP).bFloat
    
    ' fB>fG>=fR
    ElseIf (fB > fG) And (fG >= fR) Then
        '(1-fB)·[nR;nG;nB] + (fB-fG)·[nR;nG;nB+1]  +(fG-fR)·[nR;nG+1;nB+1] + (fR)·[nR+1;nG+1;nB+1]
        dstRFloat = (1! - fB) * m_lut3D(nR, nG, nB).rFloat + (fB - fG) * m_lut3D(nR, nG, nBP).rFloat + (fG - fR) * m_lut3D(nR, nGP, nBP).rFloat + fR * m_lut3D(nRP, nGP, nBP).rFloat
        dstGFloat = (1! - fB) * m_lut3D(nR, nG, nB).gFloat + (fB - fG) * m_lut3D(nR, nG, nBP).gFloat + (fG - fR) * m_lut3D(nR, nGP, nBP).gFloat + fR * m_lut3D(nRP, nGP, nBP).gFloat
        dstBFloat = (1! - fB) * m_lut3D(nR, nG, nB).bFloat + (fB - fG) * m_lut3D(nR, nG, nBP).bFloat + (fG - fR) * m_lut3D(nR, nGP, nBP).bFloat + fR * m_lut3D(nRP, nGP, nBP).bFloat
    
    ' fR>=fG>fB
    ElseIf (fR >= fG) And (fG > fB) Then
        '  (1-fR)·[nR;nG;nB] + (fR-fG)·[nR+1;nG;nB]  +(fG-fB)·[nR+1;nG+1;nB] + (fB)·[nR+1;nG+1;nB+1]
        dstRFloat = (1! - fR) * m_lut3D(nR, nG, nB).rFloat + (fR - fG) * m_lut3D(nRP, nG, nB).rFloat + (fG - fB) * m_lut3D(nRP, nGP, nB).rFloat + fB * m_lut3D(nRP, nGP, nBP).rFloat
        dstGFloat = (1! - fR) * m_lut3D(nR, nG, nB).gFloat + (fR - fG) * m_lut3D(nRP, nG, nB).gFloat + (fG - fB) * m_lut3D(nRP, nGP, nB).gFloat + fB * m_lut3D(nRP, nGP, nBP).gFloat
        dstBFloat = (1! - fR) * m_lut3D(nR, nG, nB).bFloat + (fR - fG) * m_lut3D(nRP, nG, nB).bFloat + (fG - fB) * m_lut3D(nRP, nGP, nB).bFloat + fB * m_lut3D(nRP, nGP, nBP).bFloat
    
    ' fG>fR>=fB
    ElseIf (fG > fR) And (fR >= fB) Then
        '  (1-fG)·[nR;nG;nB] + (fG-fR)·[nR;nG+1;nB]  +(fR-fB)·[nR+1;nG+1;nB] + (fB)·[nR+1;nG+1;nB+1]
        dstRFloat = (1! - fG) * m_lut3D(nR, nG, nB).rFloat + (fG - fR) * m_lut3D(nR, nGP, nB).rFloat + (fR - fB) * m_lut3D(nRP, nGP, nB).rFloat + fB * m_lut3D(nRP, nGP, nBP).rFloat
        dstGFloat = (1! - fG) * m_lut3D(nR, nG, nB).gFloat + (fG - fR) * m_lut3D(nR, nGP, nB).gFloat + (fR - fB) * m_lut3D(nRP, nGP, nB).gFloat + fB * m_lut3D(nRP, nGP, nBP).gFloat
        dstBFloat = (1! - fG) * m_lut3D(nR, nG, nB).bFloat + (fG - fR) * m_lut3D(nR, nGP, nB).bFloat + (fR - fB) * m_lut3D(nRP, nGP, nB).bFloat + fB * m_lut3D(nRP, nGP, nBP).bFloat
    
    ' fR>=fB>=fG
    Else
        '  (1-fR)·[nR;nG;nB] + (fR-fB)·[nR+1;nG;nB]  +(fB-fG)·[nR+1;nG;nB+1] + (fG)·[nR+1;nG+1;nB+1]
        dstRFloat = (1! - fR) * m_lut3D(nR, nG, nB).rFloat + (fR - fB) * m_lut3D(nRP, nG, nB).rFloat + (fB - fG) * m_lut3D(nRP, nG, nBP).rFloat + fG * m_lut3D(nRP, nGP, nBP).rFloat
        dstGFloat = (1! - fR) * m_lut3D(nR, nG, nB).gFloat + (fR - fB) * m_lut3D(nRP, nG, nB).gFloat + (fB - fG) * m_lut3D(nRP, nG, nBP).gFloat + fG * m_lut3D(nRP, nGP, nBP).gFloat
        dstBFloat = (1! - fR) * m_lut3D(nR, nG, nB).bFloat + (fR - fB) * m_lut3D(nRP, nG, nB).bFloat + (fB - fG) * m_lut3D(nRP, nG, nBP).bFloat + fG * m_lut3D(nRP, nGP, nBP).bFloat
    End If
    
    'Failsafe check and clamping
    If (dstRFloat < 0!) Then dstRFloat = 0!
    If (dstRFloat > 1!) Then dstRFloat = 1!
    If (dstGFloat < 0!) Then dstGFloat = 0!
    If (dstGFloat > 1!) Then dstGFloat = 1!
    If (dstBFloat < 0!) Then dstBFloat = 0!
    If (dstBFloat > 1!) Then dstBFloat = 1!
    dstR = Int(dstRFloat * 255! + 0.5!)
    dstG = Int(dstGFloat * 255! + 0.5!)
    dstB = Int(dstBFloat * 255! + 0.5!)
    
End Sub

Friend Sub ResetLUT()
    m_LoadSuccessful = False
    m_LutSize = 0
    m_MaxR = 1!: m_MaxG = 1!: m_MaxB = 1!
    m_MinR = 0!: m_MinG = 0!: m_MinB = 0!
    m_lutPath = vbNullString
    m_subLutPath = vbNullString
End Sub

'Serialize the current LUT's contents to a byte array.  Returns the size of the stored data, UNCOMPRESSED.
' (If you choose to compress the data, you ALSO need to retrieve and store the dstUncompressedSize value,
' obviously!)
' IMPORTANTLY: the UBound() of the returned bytes are unlikely to match the actual size of the returned data,
' because the array is over-sized to allow for more efficient compression.
Friend Function Serialize_ToBytes(ByRef dstBytes() As Byte, ByRef dstUncompressedSize As Long, ByVal cmpFormat As PD_CompressionFormat, Optional ByVal compressionLevel As Long = -1) As Long
    
    If (Not m_LoadSuccessful) Then Exit Function
    
    'As always, pdStream is the solution for all serialization needs
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_MemoryBacked, PD_SA_ReadWrite) Then
    
        'Start with a magic number, then basic header bits
        cStream.WriteString_ASCII PD_LUT_MAGIC_NUMBER
        
        'LUT path
        cStream.WriteString_UTF8 m_lutPath, True
        
        'LUT size (loop size can be inferred at read-time)
        cStream.WriteLong m_LutSize
        
        'Max/min values (range can be inferred at read-time)
        cStream.WriteFloat m_MinR
        cStream.WriteFloat m_MinG
        cStream.WriteFloat m_MinB
        cStream.WriteFloat m_MaxR
        cStream.WriteFloat m_MaxG
        cStream.WriteFloat m_MaxB
        
        'All that's left is the float data.  Swizzling the float data by byte can produce
        ' much better compression ratios, at some cost to performance.
        '
        'To make data retrievable between sessions, regardless of our behavior, write a swizzle flag
        If DEBUG_SHUFFLE_FLOATS Then cStream.WriteByte 1 Else cStream.WriteByte 0
        
        If DEBUG_SHUFFLE_FLOATS Then
            Dim tmpBytes() As Byte
            ReDim tmpBytes(0 To 4 * 3 * m_LutSize * m_LutSize * m_LutSize - 1) As Byte
            VBHacks.ShuffleBytes_4 VarPtr(m_lut3D(0, 0, 0)), VarPtr(tmpBytes(0)), 3 * m_LutSize * m_LutSize * m_LutSize
            cStream.WriteBytesFromPointer VarPtr(tmpBytes(0)), UBound(tmpBytes) + 1
        Else
            cStream.WriteBytesFromPointer VarPtr(m_lut3D(0, 0, 0)), 4 * 3 * m_LutSize * m_LutSize * m_LutSize
        End If
        
        'That's everything!  Note the uncompressed size
        dstUncompressedSize = cStream.GetStreamSize()
        
        'Figure out export size
        Dim dstArraySize As Long
        If (cmpFormat = cf_None) Then
            dstArraySize = dstUncompressedSize
        Else
            dstArraySize = Compression.GetWorstCaseSize(dstUncompressedSize, cmpFormat, compressionLevel)
        End If
        
        'Prep the array, while attempting to reuse an existing allocation (if any)
        If VBHacks.IsArrayInitialized(dstBytes) Then
            If (UBound(dstBytes) < dstArraySize - 1) Then ReDim dstBytes(0 To dstArraySize - 1) As Byte
        Else
            ReDim dstBytes(0 To dstArraySize - 1) As Byte
        End If
        
        'Compress!
        If (cmpFormat = cf_None) Then
            cStream.SetPosition 0
            cStream.ReadBytesToBarePointer VarPtr(dstBytes(0)), dstUncompressedSize
            Serialize_ToBytes = dstUncompressedSize
        Else
        
            Dim compressedSize As Long
            compressedSize = dstArraySize
            
            If Compression.CompressPtrToPtr(VarPtr(dstBytes(0)), compressedSize, cStream.Peek_PointerOnly(0, dstUncompressedSize), dstUncompressedSize, cmpFormat, compressionLevel) Then
                'Debug.Print "Compressed LUT during serialization: " & cStream.GetStreamSize() & " to " & compressedSize & " bytes"
                Serialize_ToBytes = compressedSize
            Else
                InternalProblem "Serialize_ToBytes", "compression failure"
                Serialize_ToBytes = 0
            End If
            
        End If
        
    End If
    
End Function

'Re-create a LUT from previously serialized bytes.  Returns TRUE if un-serialization appears to be successful.
Friend Function Serialize_FromBytes(ByRef srcBytes() As Byte, ByVal currentSizeInBytes As Long, ByVal uncompressedSizeInBytes As Long, ByVal cmpFormat As PD_CompressionFormat) As Boolean
    
    'Reset everything we've currently got stored
    Me.ResetLUT
    
    'As always, pdStream is the solution for all serialization needs
    Dim cStream As pdStream
    Set cStream = New pdStream
    
    'If the data is compressed, we need to uncompress it before reading.
    Dim streamOK As Boolean
    If (cmpFormat = cf_None) Then
        
        'Wrap the stream around the data as-is
        streamOK = cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, startingBufferSize:=currentSizeInBytes, baseFilePointerOffset:=VarPtr(srcBytes(0)))
    
    Else
    
        'Decompress the data directly into the stream object
        streamOK = cStream.StartStream(PD_SM_MemoryBacked, PD_SA_ReadWrite, startingBufferSize:=uncompressedSizeInBytes)
        cStream.SetSizeExternally uncompressedSizeInBytes
        If streamOK Then streamOK = Compression.DecompressPtrToPtr(cStream.Peek_PointerOnly(0, uncompressedSizeInBytes), uncompressedSizeInBytes, VarPtr(srcBytes(0)), currentSizeInBytes, cmpFormat)
        If streamOK Then cStream.SetPosition 0, FILE_BEGIN
        
    End If
    
    If (Not streamOK) Then
        InternalProblem "Serialize_FromByte", "couldn't initialize stream"
        Exit Function
    End If
    
    'Validate the magic number
    Dim testString As String
    testString = cStream.ReadString_ASCII(8)
    If (testString <> PD_LUT_MAGIC_NUMBER) Then
        InternalProblem "Serialize_FromByte", "magic number doesn't match: " & testString
        Exit Function
    End If
    
    'Retrieve LUT path
    m_lutPath = cStream.ReadString_UTF8(0, True)
    
    'Load and validate LUT size
    m_LutSize = cStream.ReadLong()
    If (m_LutSize <= 0) Then
        InternalProblem "Serialize_FromByte", "bad lut size: " & m_LutSize
        Exit Function
    End If
        
    m_LutLoopSize = m_LutSize - 1
        
    'Max/min values, then infer range accordingly
    m_MinR = cStream.ReadFloat()
    m_MinG = cStream.ReadFloat()
    m_MinB = cStream.ReadFloat()
    m_MaxR = cStream.ReadFloat()
    m_MaxG = cStream.ReadFloat()
    m_MaxB = cStream.ReadFloat()
    
    m_RangeR = m_MaxR - m_MinR
    m_RangeG = m_MaxG - m_MinG
    m_RangeB = m_MaxB - m_MinB
    
    'All that's left is float data.  Prepare the array then copy the data directly into place.
    ReDim m_lut3D(0 To m_LutLoopSize, 0 To m_LutLoopSize, 0 To m_LutLoopSize) As FloatColor
    
    'Float data may be swizzled for improved compression performance.  Retrieve the swizzle flag first
    If (cStream.ReadByte() = 1) Then
        Dim tmpBytes() As Byte
        ReDim tmpBytes(0 To 4 * 3 * m_LutSize * m_LutSize * m_LutSize - 1) As Byte
        cStream.ReadBytesToBarePointer VarPtr(tmpBytes(0)), UBound(tmpBytes) + 1
        VBHacks.UnshuffleBytes_4 VarPtr(tmpBytes(0)), VarPtr(m_lut3D(0, 0, 0)), 3 * m_LutSize * m_LutSize * m_LutSize
    Else
        cStream.ReadBytesToBarePointer VarPtr(m_lut3D(0, 0, 0)), 4 * 3 * m_LutSize * m_LutSize * m_LutSize
    End If
    
    m_LoadSuccessful = True
    Serialize_FromBytes = m_LoadSuccessful

End Function

Private Sub InternalProblem(ByRef srcFunction As String, ByRef errMsg As String)
    PDDebug.LogAction "WARNING!  pdLUT3D." & srcFunction & "() returned an error: " & errMsg
End Sub
