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
'Copyright 2020-2026 by Tanner Helland
'Created: 23/October/20
'Last updated: 01/June/22
'Last update: new support for resizing LUT dimensions on-the-fly
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
' (links good as of October 2020):
'
' - Version 1.0 of Adobe's .CUBE spec:
'    https://wwwimages2.adobe.com/content/dam/acom/en/products/speedgrade/cc/pdfs/cube-lut-specification-1.0.pdf
' - Technicolor presentation (via NVidia) on tetrahedral interpolation:
'    https://www.nvidia.com/content/GTC/posters/2010/V01-Real-Time-Color-Space-Conversion-for-High-Resolution-Video.pdf
' - 2009 version of Autodesk's 3DL spec:
'    https://download.autodesk.com/us/systemdocs/pdf/lustre_color_management_user_guide.pdf
'
'This class has been heavily optimized for speed, which means it has many dependencies on other
' PhotoDemon classes.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Detailed debug reports can be useful for profiling and/or troubleshooting
Private Const DEBUG_VERBOSE As Boolean = True

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

'When creating our own LUTs from before-and-after images, we have to perform heuristics to try and
' "guess" at what the user did to the image.
Private Type PD_LUTConstructor
    origColor As RGBQuad
    newColor As RGBQuad
    fxColor As RGBQuad
    finalColor As RGBQuad
    posInImage As PointLong
    origColorDistance As Long
End Type

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

Public Function BuildLUTFromTwoDIBs(ByRef baseDIB As pdDIB, ByRef modifiedDIB As pdDIB, ByVal numSamplingPoints As Long, Optional ByVal showProgressBar As Boolean = False) As Boolean
    
    Const FUNC_NAME As String = "BuildLUTFromTwoDIBs"
    BuildLUTFromTwoDIBs = True
    
    'Validate sampling points
    If (numSamplingPoints < 0) Or (numSamplingPoints > 64) Then
        InternalProblem FUNC_NAME, "bad sampling points"
        BuildLUTFromTwoDIBs = False
        Exit Function
    End If
    
    If showProgressBar Then
        ProgressBars.SetProgBarMax 10
        ProgressBars.SetProgBarVal 0
        Message "Analyzing image..."
    End If
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Initializing LUT creator..."
    
    'Initialize the module-level LUT to the correct grid size, and fill it with identity values
    m_LutSize = numSamplingPoints
    m_LutLoopSize = numSamplingPoints - 1
    ReDim m_lut3D(0 To m_LutLoopSize, 0 To m_LutLoopSize, 0 To m_LutLoopSize) As FloatColor
    
    Dim r As Long, g As Long, b As Long
    Dim newR As Single, newG As Single, newB As Single
    
    For r = 0 To m_LutLoopSize
        newR = r / m_LutLoopSize
    For g = 0 To m_LutLoopSize
        newG = g / m_LutLoopSize
    For b = 0 To m_LutLoopSize
        newB = b / m_LutLoopSize
        With m_lut3D(r, g, b)
            .rFloat = newR
            .gFloat = newG
            .bFloat = newB
        End With
    Next b
    Next g
    Next r
    
    If showProgressBar Then ProgressBars.SetProgBarVal 1
    
    'We now have an identity LUT at the correct size, and it has been initialized to an identity transform.
    
    'Our next goal is to iterate the passed images, and for each pixel, find its nearest point in the LUT.
    ' This will produce an array of "best-matched" points, from which we can roughly construct a new LUT.
    
    'Start by creating a LUT constructor array; this will track "best-match" points, colors, and positions
    ' from the source image.
    Dim curAnalysis() As PD_LUTConstructor
    ReDim curAnalysis(0 To m_LutLoopSize, 0 To m_LutLoopSize, 0 To m_LutLoopSize) As PD_LUTConstructor
    
    For r = 0 To m_LutLoopSize
    For g = 0 To m_LutLoopSize
    For b = 0 To m_LutLoopSize
        With curAnalysis(r, g, b)
            .origColorDistance = LONG_MAX
            .origColor.Red = Int((r / m_LutLoopSize) * 255 + 0.5)
            .origColor.Green = Int((g / m_LutLoopSize) * 255 + 0.5)
            .origColor.Blue = Int((b / m_LutLoopSize) * 255 + 0.5)
        End With
    Next b
    Next g
    Next r
    
    Dim curTime As Currency
    If DEBUG_VERBOSE Then
        PDDebug.LogAction "LUT creation took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime curTime
    End If
    
    If showProgressBar Then ProgressBars.SetProgBarVal 2
    
    'New approach!  Scan the base image and build a list of colors from it - and importantly, store each color alongside
    ' its position in the image.
    Dim cHash As pdHashCoord
    Set cHash = New pdHashCoord
    
    'Scan the *base* image and for each color, convert it to its nearest LUT point, compare its distance
    ' to the current best-match and update accordingly.
    Dim pxBase() As Long, pxSA As SafeArray1D
    
    Dim x As Long, y As Long, xMax As Long, yMax As Long
    xMax = baseDIB.GetDIBWidth - 1
    yMax = baseDIB.GetDIBHeight - 1
    
    Dim tmpPoint As PointLong, tmpColor As RGBQuad, tmpLong As Long
    
    For y = 0 To yMax
        tmpPoint.y = y
        baseDIB.WrapLongArrayAroundScanline pxBase, pxSA, y
    For x = 0 To xMax
        
        tmpPoint.x = x
        
        'Force alpha to 255; LUTs don't consider transparency
        tmpLong = pxBase(x) Or &HFF000000
        
        'Add the color from this location to the hash table.  The hash table will automatically remove duplicates,
        ' and it will store coordinates alongside colors, giving us a 1:1 mapping between the two.
        cHash.AddColor tmpLong, tmpPoint
        
    Next x
    Next y
    
    baseDIB.UnwrapLongArrayFromDIB pxBase
    If showProgressBar Then ProgressBars.SetProgBarVal 3
    
    'With a hash-table assembled, we now need to build a KD-tree of the final color list.  We can traverse this tree
    ' to find best-match colors for each LUT entry.
    
    'Start by pulling a list of unique colors from the hash table.
    Dim listOfColors() As RGBQuad, numOfColors As Long
    numOfColors = cHash.GetUniqueColors(listOfColors)
    
    If DEBUG_VERBOSE Then
        PDDebug.LogAction "Hash table (and final color list retrieval) took " & VBHacks.GetTimeDiffNowAsString(curTime)
        PDDebug.LogAction "(Hash table has " & numOfColors & " colors.)"
        VBHacks.GetHighResTime curTime
    End If
    If showProgressBar Then ProgressBars.SetProgBarVal 4
    
    'Add each color from the list to a KD-tree; we will momentarily use this tree to retrieve best-fit colors
    Dim cTree As pdKDTreeArray
    Set cTree = New pdKDTreeArray
    cTree.BuildTree listOfColors, 0, numOfColors - 1
    
    If DEBUG_VERBOSE Then
        PDDebug.LogAction "KD-tree construction took " & VBHacks.GetTimeDiffNowAsString(curTime)
        VBHacks.GetHighResTime curTime
    End If
    
    If showProgressBar Then
        ProgressBars.SetProgBarVal 5
        Message "Building color lookup..."
    End If
    
    'With a KD-tree assembled and a corresponding color-to-coordinate table, we can finally start matching LUT colors.
    
    'The way we're gonna do this is iterate through the LUT, and for each point, find the closest matching color
    ' in the original image (using the KD-tree).  From these "best-fit" colors, we'll construct a "best-fit" LUT
    ' approximation that describes the changes from the base image to the modified image across the full spectrum
    ' covered by this photo.
    For r = 0 To m_LutLoopSize
    For g = 0 To m_LutLoopSize
    For b = 0 To m_LutLoopSize
        
        'Retrieve the matching color at this point, force full opacity, then store the result
        tmpColor = cTree.GetNearestColor(curAnalysis(r, g, b).origColor)
        tmpColor.Alpha = 255
        curAnalysis(r, g, b).newColor = tmpColor
        
    Next b
    Next g
    Next r
    
    'We are done with the tree class; kill it to free resources
    Set cTree = Nothing
    
    If DEBUG_VERBOSE Then
        PDDebug.LogAction "KD-tree matching took " & VBHacks.GetTimeDiffNowAsString(curTime)
        VBHacks.GetHighResTime curTime
    End If
    If showProgressBar Then ProgressBars.SetProgBarVal 6
    
    'We now have a "best-match" color for every LUT entry *using only colors that existed in the original,
    ' unmodified target image*.
    '
    'That last point is important because no (real-world) image contains colors from the entire color spectrum.
    ' Because of this, some regions of our LUT will have been matched extremely well, while others will be
    ' matched very poorly.  That's OK - we'll deal with it as best we can in a moment.
    '
    'For now, however, we need to retrieve the *modified* version of each matched *original* color in the image.
    ' We can do this using our original hash table (which cleverly stored color and coordinate data alongside
    ' each other).  This will give us a target to use for interpolating new color values on colors that were
    ' matched poorly.
    For r = 0 To m_LutLoopSize
    For g = 0 To m_LutLoopSize
    For b = 0 To m_LutLoopSize
        
        With curAnalysis(r, g, b)
            GetMem4 VarPtr(.newColor), tmpLong
            If cHash.GetCoordForColor(tmpLong, tmpPoint) Then
                .posInImage = tmpPoint
            Else
                'Never expected; colors must always match!
                If DEBUG_VERBOSE Then InternalProblem FUNC_NAME, "color doesn't exist in hash table"
            End If
        End With
    
    Next b
    Next g
    Next r
    
    'We are done with the hash table; kill it to free up resources
    Set cHash = Nothing
    
    If DEBUG_VERBOSE Then
        PDDebug.LogAction "Hash table coordinate retrieval took " & VBHacks.GetTimeDiffNowAsString(curTime)
        VBHacks.GetHighResTime curTime
    End If
    If showProgressBar Then ProgressBars.SetProgBarVal 7
    
    Dim pxFinal() As RGBQuad, pxSA2D As SafeArray2D
    modifiedDIB.WrapRGBQuadArrayAroundDIB pxFinal, pxSA2D
    
    For r = 0 To m_LutLoopSize
    For g = 0 To m_LutLoopSize
    For b = 0 To m_LutLoopSize
        With curAnalysis(r, g, b)
            .fxColor = pxFinal(.posInImage.x, .posInImage.y)
        End With
    Next b
    Next g
    Next r
    
    If DEBUG_VERBOSE Then
        PDDebug.LogAction "Modified color retrievals took " & VBHacks.GetTimeDiffNowAsString(curTime)
        VBHacks.GetHighResTime curTime
    End If
    If showProgressBar Then ProgressBars.SetProgBarVal 8
    
    'Free unsafe array wrapper and note that we no longer need to touch either DIB
    modifiedDIB.UnwrapRGBQuadArrayFromDIB pxFinal
    
    'Every entry in the LUT now has a best-match original color, the corresponding "final" color (with all
    ' adjustments+effects applied), and a coordinate where said color appears in the image (which probably
    ' isn't useful at this point so maybe we don't need to store it at all?  TODO!)
    
    'We now need to scale those best-match colors to reflect the difference between the color the LUT represents
    ' (which is some fixed-interval value like [0, 0, 0], [0, 0, 15], [0, 0, 31] etc) and whatever the
    ' best-match color for that LUT value is (like maybe [0, 0, 31] got matched to [5, 12, 16]).  We do this by
    ' calculating the difference between each LUT value and its best-matched color, then applying that same
    ' scale to the *new* color - this is just cheap 1D interpolation on each axis and is not necessarily ideal,
    ' but it's very fast and cheap to implement.  (We could always do a second pass over the data, and for
    ' LUT entries that were poorly matched - perhaps outside some threshold - we could average them against
    ' their neighbors to smooth out the LUT mesh - exploring this is TODO.)
    '
    'Also TODO: vectors are almost certainly a smarter approach here - calculate a 3D vector from the source
    ' color to the final color, then scale that vector according to its position in the overall 3D cube when
    ' applying it to the LUT color.  Researching this alternative implementation is TODO.
    Dim fxScale As Single, tmpFloat As Single
    
    For r = 0 To m_LutLoopSize
    For g = 0 To m_LutLoopSize
    For b = 0 To m_LutLoopSize
        
        With curAnalysis(r, g, b)
            
            'Remap the effect color linearly.  (We could presumably do more intense/accurate interpolation,
            ' but since this whole approach is hackish I'm not worried about linear quality implications.)
            If (.newColor.Blue > 0) Then
                fxScale = (.fxColor.Blue / .newColor.Blue)
            Else
                fxScale = 0!
            End If
            
            'Scale the "idealized" LUT value at this point
            tmpFloat = m_lut3D(r, g, b).bFloat * fxScale
            If (tmpFloat > 1!) Then tmpFloat = 1!
            m_lut3D(r, g, b).bFloat = tmpFloat
            
            If (.newColor.Green > 0) Then
                fxScale = (.fxColor.Green / .newColor.Green)
            Else
                fxScale = 0!
            End If
            
            tmpFloat = m_lut3D(r, g, b).gFloat * fxScale
            If (tmpFloat > 1!) Then tmpFloat = 1!
            m_lut3D(r, g, b).gFloat = tmpFloat
            
            If (.newColor.Red > 0) Then
                fxScale = (.fxColor.Red / .newColor.Red)
            Else
                fxScale = 0!
            End If
            
            tmpFloat = m_lut3D(r, g, b).rFloat * fxScale
            If (tmpFloat > 1!) Then tmpFloat = 1!
            m_lut3D(r, g, b).rFloat = tmpFloat
            
            'With a baseline color generated, we now want to cover some special cases.
            
            'First, look for grayscale conversions.  We can detect this if the R/G/B values of the target color
            ' are all identical.  (Various corrections either do not apply to grayscale pixels - e.g. "vibrance" -
            ' or they forcibly convert certain tones to grayscale - e.g. "selective colorize".  We can cover these
            ' cases nicely with a custom check.)
            If (.fxColor.Red = .fxColor.Green) And (.fxColor.Green = .fxColor.Blue) Then
                
                'Average the new colors we calculated
                tmpFloat = (m_lut3D(r, g, b).rFloat + m_lut3D(r, g, b).gFloat + m_lut3D(r, g, b).bFloat) * 0.333333333333333
                m_lut3D(r, g, b).rFloat = tmpFloat
                m_lut3D(r, g, b).gFloat = tmpFloat
                m_lut3D(r, g, b).bFloat = tmpFloat
                
            End If
            
            'Next, look for inversion.  This is a challenging one to detect via linear interpolation,
            ' but it's easily detected using the source colors.
            If (.fxColor.Red = 255 - .newColor.Red) And (.fxColor.Green = 255 - .newColor.Green) And (.fxColor.Blue = 255 - .newColor.Blue) Then
                
                'Invert this color accordingly
                m_lut3D(r, g, b).rFloat = (m_LutLoopSize - r) / m_LutLoopSize
                m_lut3D(r, g, b).gFloat = (m_LutLoopSize - g) / m_LutLoopSize
                m_lut3D(r, g, b).bFloat = (m_LutLoopSize - b) / m_LutLoopSize
                
            End If
            
        End With
        
    Next b
    Next g
    Next r
    
    If DEBUG_VERBOSE Then
        PDDebug.LogAction "Final scaling and interpolation took " & VBHacks.GetTimeDiffNowAsString(curTime)
        VBHacks.GetHighResTime curTime
    End If
    If showProgressBar Then ProgressBars.SetProgBarVal 9
    
    'For now, let's just return TRUE so I can work on export features.
    BuildLUTFromTwoDIBs = True
    
End Function

'Display PD's generic 3D LUT load dialog.  All supported LUT filetypes are available to the user.
' Note that this function returns TRUE if the user selects one or more files, FALSE if they cancel.
' This function DOES NOT attempt to load or validate target files - that's the caller's responsibility.
Public Function DisplayLUTLoadDialog(ByRef srcFilename As String, ByRef dstFilenames As pdStringStack) As Boolean
    
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
    
    If openDialog.GetOpenFileNames_AsStringStack(dstFilenames, sFile, vbNullString, True, cdFilter.ToString(), 1, UserPrefs.GetLUTPath(False), cdTitle, , GetModalOwner().hWnd) Then
    
        'By design, we don't perform any validation here.  Let the caller validate the file as much (or as little)
        ' as they require.
        DisplayLUTLoadDialog = (Not dstFilenames Is Nothing)
        If DisplayLUTLoadDialog Then DisplayLUTLoadDialog = (dstFilenames.GetNumOfStrings > 0)
        
        'The dialog was successful.  Save this path to the user's pref file.
        If DisplayLUTLoadDialog Then UserPrefs.SetLUTPath Files.FileGetPath(dstFilenames.GetString(0))
        
    End If
    
    'Re-enable user input
    Interface.EnableUserInput
    
End Function

Friend Function GetLUTPath() As String
    GetLUTPath = m_lutPath
End Function

Friend Function GetLUTSize() As Long
    GetLUTSize = m_LutSize
End Function

'Only some files (.look format) support the concept of subpaths
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
                            m_LutSize = Val(Strings.RightByChar(curLine, spaceChar))
                            m_LutLoopSize = m_LutSize - 1
                            
                        'Minimum values, typically 0.0
                        ElseIf Strings.LeftMatches(curLine, "DOMAIN_MIN", True) Then
                            If firstLineHit Then InternalProblem "LoadLUTFromFile", "bad line order"
                            
                            'Retrieve all three values
                            Set wordsInLine = Strings.GetListOfWordsFromString(curLine)
                            m_MinR = Val(wordsInLine.GetString(1))
                            m_MinG = Val(wordsInLine.GetString(2))
                            m_MinB = Val(wordsInLine.GetString(3))
                            
                        ElseIf Strings.LeftMatches(curLine, "DOMAIN_MAX", True) Then
                            If firstLineHit Then InternalProblem "LoadLUTFromFile", "bad line order"
                            
                            Set wordsInLine = Strings.GetListOfWordsFromString(curLine)
                            m_MaxR = Val(wordsInLine.GetString(1))
                            m_MaxG = Val(wordsInLine.GetString(2))
                            m_MaxB = Val(wordsInLine.GetString(3))
                            
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

'Passed values must be on the range [0.0, 1.0].  Returned values are currently clamped to [0.0, 1.0] as well
' (but this could easily be modified as desired).
Friend Sub GetInterpolatedRGB_Float(ByVal srcR As Single, ByVal srcG As Single, ByVal srcB As Single, ByRef dstR As Single, ByRef dstG As Single, ByRef dstB As Single)

    'Use the source RGB values as indices into the central lookup table
    Dim rIndex As Single, gIndex As Single, bIndex As Single
    rIndex = (srcR * m_LutLoopSize)
    gIndex = (srcG * m_LutLoopSize)
    bIndex = (srcB * m_LutLoopSize)
    
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
    
    'Branches enable faster conversions
    If (fG >= fB) And (fB >= fR) Then
        dstR = (1! - fG) * m_lut3D(nR, nG, nB).rFloat + (fG - fB) * m_lut3D(nR, nGP, nB).rFloat + (fB - fR) * m_lut3D(nR, nGP, nBP).rFloat + fR * m_lut3D(nRP, nGP, nBP).rFloat
        dstG = (1! - fG) * m_lut3D(nR, nG, nB).gFloat + (fG - fB) * m_lut3D(nR, nGP, nB).gFloat + (fB - fR) * m_lut3D(nR, nGP, nBP).gFloat + fR * m_lut3D(nRP, nGP, nBP).gFloat
        dstB = (1! - fG) * m_lut3D(nR, nG, nB).bFloat + (fG - fB) * m_lut3D(nR, nGP, nB).bFloat + (fB - fR) * m_lut3D(nR, nGP, nBP).bFloat + fR * m_lut3D(nRP, nGP, nBP).bFloat
    ElseIf (fB > fR) And (fR > fG) Then
        dstR = (1! - fB) * m_lut3D(nR, nG, nB).rFloat + (fB - fR) * m_lut3D(nR, nG, nBP).rFloat + (fR - fG) * m_lut3D(nRP, nG, nBP).rFloat + fG * m_lut3D(nRP, nGP, nBP).rFloat
        dstG = (1! - fB) * m_lut3D(nR, nG, nB).gFloat + (fB - fR) * m_lut3D(nR, nG, nBP).gFloat + (fR - fG) * m_lut3D(nRP, nG, nBP).gFloat + fG * m_lut3D(nRP, nGP, nBP).gFloat
        dstB = (1! - fB) * m_lut3D(nR, nG, nB).bFloat + (fB - fR) * m_lut3D(nR, nG, nBP).bFloat + (fR - fG) * m_lut3D(nRP, nG, nBP).bFloat + fG * m_lut3D(nRP, nGP, nBP).bFloat
    ElseIf (fB > fG) And (fG >= fR) Then
        dstR = (1! - fB) * m_lut3D(nR, nG, nB).rFloat + (fB - fG) * m_lut3D(nR, nG, nBP).rFloat + (fG - fR) * m_lut3D(nR, nGP, nBP).rFloat + fR * m_lut3D(nRP, nGP, nBP).rFloat
        dstG = (1! - fB) * m_lut3D(nR, nG, nB).gFloat + (fB - fG) * m_lut3D(nR, nG, nBP).gFloat + (fG - fR) * m_lut3D(nR, nGP, nBP).gFloat + fR * m_lut3D(nRP, nGP, nBP).gFloat
        dstB = (1! - fB) * m_lut3D(nR, nG, nB).bFloat + (fB - fG) * m_lut3D(nR, nG, nBP).bFloat + (fG - fR) * m_lut3D(nR, nGP, nBP).bFloat + fR * m_lut3D(nRP, nGP, nBP).bFloat
    ElseIf (fR >= fG) And (fG > fB) Then
        dstR = (1! - fR) * m_lut3D(nR, nG, nB).rFloat + (fR - fG) * m_lut3D(nRP, nG, nB).rFloat + (fG - fB) * m_lut3D(nRP, nGP, nB).rFloat + fB * m_lut3D(nRP, nGP, nBP).rFloat
        dstG = (1! - fR) * m_lut3D(nR, nG, nB).gFloat + (fR - fG) * m_lut3D(nRP, nG, nB).gFloat + (fG - fB) * m_lut3D(nRP, nGP, nB).gFloat + fB * m_lut3D(nRP, nGP, nBP).gFloat
        dstB = (1! - fR) * m_lut3D(nR, nG, nB).bFloat + (fR - fG) * m_lut3D(nRP, nG, nB).bFloat + (fG - fB) * m_lut3D(nRP, nGP, nB).bFloat + fB * m_lut3D(nRP, nGP, nBP).bFloat
    ElseIf (fG > fR) And (fR >= fB) Then
        dstR = (1! - fG) * m_lut3D(nR, nG, nB).rFloat + (fG - fR) * m_lut3D(nR, nGP, nB).rFloat + (fR - fB) * m_lut3D(nRP, nGP, nB).rFloat + fB * m_lut3D(nRP, nGP, nBP).rFloat
        dstG = (1! - fG) * m_lut3D(nR, nG, nB).gFloat + (fG - fR) * m_lut3D(nR, nGP, nB).gFloat + (fR - fB) * m_lut3D(nRP, nGP, nB).gFloat + fB * m_lut3D(nRP, nGP, nBP).gFloat
        dstB = (1! - fG) * m_lut3D(nR, nG, nB).bFloat + (fG - fR) * m_lut3D(nR, nGP, nB).bFloat + (fR - fB) * m_lut3D(nRP, nGP, nB).bFloat + fB * m_lut3D(nRP, nGP, nBP).bFloat
    Else
        dstR = (1! - fR) * m_lut3D(nR, nG, nB).rFloat + (fR - fB) * m_lut3D(nRP, nG, nB).rFloat + (fB - fG) * m_lut3D(nRP, nG, nBP).rFloat + fG * m_lut3D(nRP, nGP, nBP).rFloat
        dstG = (1! - fR) * m_lut3D(nR, nG, nB).gFloat + (fR - fB) * m_lut3D(nRP, nG, nB).gFloat + (fB - fG) * m_lut3D(nRP, nG, nBP).gFloat + fG * m_lut3D(nRP, nGP, nBP).gFloat
        dstB = (1! - fR) * m_lut3D(nR, nG, nB).bFloat + (fR - fB) * m_lut3D(nRP, nG, nB).bFloat + (fB - fG) * m_lut3D(nRP, nG, nBP).bFloat + fG * m_lut3D(nRP, nGP, nBP).bFloat
    End If
    
    'Clamping could be optional for floating-point values, but we currently desire this behavior in PD.
    If (dstR < 0!) Then dstR = 0!
    If (dstR > 1!) Then dstR = 1!
    If (dstG < 0!) Then dstG = 0!
    If (dstG > 1!) Then dstG = 1!
    If (dstB < 0!) Then dstB = 0!
    If (dstB > 1!) Then dstB = 1!
    
End Sub

'Passed values should be on the range [0, 255]
Friend Sub GetInterpolatedRGB_Int(ByVal srcR As Long, ByVal srcG As Long, ByVal srcB As Long, ByRef dstR As Long, ByRef dstG As Long, ByRef dstB As Long)

    'Use the source RGB values as indices into the central lookup table
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

'If a LUT has been loaded, you can resize it to any arbitrary size via this function.
' This can convert e.g. a 33x33x33 LUT to a 17x17x17 one.  LUT size doesn't matter much for performance,
' but reducing it can greatly reduce the on-disk *size* of a LUT.  (In the aforementioned example,
' the 17x17x17 LUT only has ~5000 entries compared to ~36000 for the 33x33x33 version.)
Friend Sub ResizeLUT(ByVal newLUTSize As Long)
    
    'Resizing a LUT is easy: just create a temporary set of floats, run the LUT filter on it, then swap the
    ' module-level values out.
    Dim newBound As Long
    newBound = newLUTSize - 1
    
    Dim newValues() As FloatColor
    ReDim newValues(0 To newLUTSize - 1, 0 To newLUTSize - 1, 0 To newLUTSize - 1) As FloatColor
    
    Dim r As Long, g As Long, b As Long
    Dim origR As Single, origG As Single, origB As Single, newR As Single, newG As Single, newB As Single
    
    Dim minR As Single, minG As Single, minB As Single
    Dim maxR As Single, maxG As Single, maxB As Single
    minR = 0!: minG = 0!: minB = 0!
    maxR = 1!: maxG = 1!: maxB = 1!
    
    For r = 0 To newBound
        origR = r / newBound
    For g = 0 To newBound
        origG = g / newBound
    For b = 0 To newBound
        origB = b / newBound
        Me.GetInterpolatedRGB_Float origR, origG, origB, newR, newG, newB
        If (newR < minR) Then minR = newR
        If (newR > maxR) Then maxR = newR
        If (newG < minG) Then minG = newG
        If (newG > maxG) Then maxG = newG
        If (newB < minB) Then minB = newB
        If (newB > maxB) Then maxB = newB
        With newValues(r, g, b)
            .rFloat = newR
            .gFloat = newG
            .bFloat = newB
        End With
    Next b
    Next g
    Next r
    
    'With new values calculated, swap out the class-level LUT with our new one.
    ReDim m_lut3D(0 To newBound, 0 To newBound, 0 To newBound) As FloatColor
    m_LutSize = newLUTSize
    m_LutLoopSize = newLUTSize - 1
    
    m_MinR = minR: m_MinG = minG: m_MinB = minB
    m_MaxR = maxR: m_MaxG = maxG: m_MaxB = maxB
    m_RangeR = m_MaxR - m_MinR: m_RangeG = m_MaxG - m_MinG: m_RangeB = m_MaxB - m_MinB
    
    For r = 0 To newBound
    For g = 0 To newBound
    For b = 0 To newBound
        m_lut3D(r, g, b) = newValues(r, g, b)
    Next b
    Next g
    Next r
    
End Sub

'Export the current LUT to a 3dl file
Friend Function SaveLUTToFile_3dl(ByRef dstFile As String, Optional ByRef copyrightText As String = vbNullString, Optional ByRef textDescription As String = vbNullString) As Boolean
    
    Const FUNC_NAME As String = "SaveLUTToFile_3dl"
    
    On Error GoTo CouldNotSaveFile
    
    'Failsafe checks for valid inputs and/or class state
    Dim cantSave As Boolean: cantSave = False
    
    If (LenB(dstFile) = 0) Then
        cantSave = True
        InternalProblem FUNC_NAME, "invalid function parameters"
    End If
    
    If (m_LutSize <= 0) Then
        cantSave = True
        InternalProblem FUNC_NAME, "no LUT"
    End If
    
    If cantSave Then
        SaveLUTToFile_3dl = False
        Exit Function
    End If
    
    'Failsafe check only; PD should enforce safe overwriting in the parent function.
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    Dim startTime As Currency
    If DEBUG_VERBOSE Then VBHacks.GetHighResTime startTime
    
    'To improve performance, write everything to a pdString instance.  Note that we'll be using
    ' Unix-style Lf line terminators (*not* CrLf) for compatibility reasons.  (Technically 3DL files
    ' should work either way, but nearly all the ones I see "in the wild" have Unix line-endings.)
    Dim cTextFile As pdString
    Set cTextFile = New pdString
    
    '3DL files support comments, and convention has you stick the app, copyright, and description
    ' on the first three lines of the file as comments, e.g. this from an actual 3DL file:
    '#Created by: Adobe Photoshop Export Color Lookup Plugin
    '#Copyright: (C) Copyright 1234 [various names here]
    '#Description: [put whatever you want here]
    cTextFile.AppendLineLf "#Created by: " & Updates.GetPhotoDemonNameAndVersion()
    If (LenB(copyrightText) <> 0) Then cTextFile.AppendLineLf "#Copyright: (c) " & Year(Now) & " " & copyrightText
    If (LenB(textDescription) <> 0) Then cTextFile.AppendLineLf "#Description: " & textDescription
    
    'After the above lines comes a fixed list of sampling intervals.  3DL files use 10-bit color defs,
    ' so these are values on the range [0, 1023].  The default number of discrete values is 17 (per the spec).
    Dim i As Long
    For i = 0 To m_LutLoopSize
        cTextFile.Append Trim$(Str$(Int((i / m_LutLoopSize) * 1023 + 0.5)))
        If (i < m_LutLoopSize) Then cTextFile.Append " "
    Next i
    
    cTextFile.AppendLineBreakLf
    
    'We now append all colors in RGB order, with the B channel incrementing fastest.
    ' (For reasons unknown to me, values are written on the range [0, 4095] despite 10-bit
    ' color being used.  I do not know the reason for this and the 3DL spec doesn't explain it -
    ' but this is how all the 3DL files that ship with Photoshop work, so we do it too.
    Dim r As Long, g As Long, b As Long
    
    For r = 0 To m_LutLoopSize
    For g = 0 To m_LutLoopSize
    For b = 0 To m_LutLoopSize
        With m_lut3D(r, g, b)
            cTextFile.AppendLineLf Trim$(Str$(Int(.rFloat * 4095 + 0.5))) & Str$(Int(.gFloat * 4095 + 0.5)) & Str$(Int(.bFloat * 4095 + 0.5))
        End With
    Next b
    Next g
    Next r
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Constructing final 3DL text took " & VBHacks.GetTimeDiffNowAsString(startTime)
    
    'All that's left is to write the table out to file!
    If DEBUG_VERBOSE Then PDDebug.LogAction "Handing final 3DL text off to Files module for export to " & dstFile
    SaveLUTToFile_3dl = Files.FileSaveAsText(cTextFile.ToString, dstFile, True, False)
    
    Exit Function
    
CouldNotSaveFile:
    InternalProblem FUNC_NAME, "VB error #" & Err.Number & ": " & Err.Description
    SaveLUTToFile_3dl = False
    
End Function

'Export the current LUT to a cube file
Friend Function SaveLUTToFile_Cube(ByRef dstFile As String, Optional ByRef copyrightText As String = vbNullString, Optional ByRef textDescription As String = vbNullString) As Boolean
    
    Const FUNC_NAME As String = "SaveLUTToFile_Cube"
    
    On Error GoTo CouldNotSaveFile
    
    'Failsafe checks for valid inputs and/or class state
    Dim cantSave As Boolean: cantSave = False
    
    If (LenB(dstFile) = 0) Then
        cantSave = True
        InternalProblem FUNC_NAME, "invalid function parameters"
    End If
    
    If (m_LutSize <= 0) Then
        cantSave = True
        InternalProblem FUNC_NAME, "no LUT"
    End If
    
    If cantSave Then
        SaveLUTToFile_Cube = False
        Exit Function
    End If
    
    'Failsafe check only; PD should enforce safe overwriting in the parent function.
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    Dim startTime As Currency
    If DEBUG_VERBOSE Then VBHacks.GetHighResTime startTime
    
    'To improve performance, write everything to a pdString instance.  Note that we'll be using
    ' Unix-style Lf line terminators (*not* CrLf) for compatibility reasons.  (Technically cube files
    ' should work either way, but nearly all the ones I see "in the wild" have Unix line-endings.)
    Dim cTextFile As pdString
    Set cTextFile = New pdString
    
    'Cube files support comments, and convention has you stick the app and copyright, and description
    ' on the first two lines of the file as comments, e.g. this from an actual cube file:
    '#Created by: Adobe Photoshop Export Color Lookup Plugin
    '#Copyright: (C) Copyright 1234 [various names here]
    cTextFile.AppendLineLf "#Created by: " & Updates.GetPhotoDemonNameAndVersion()
    If (LenB(copyrightText) <> 0) Then cTextFile.AppendLineLf "#Copyright: (c) " & Year(Now) & " " & copyrightText
    
    'LUT title is its own block, prefixed by a keyword (you'll see this pattern elsewhere in the file)
    If (LenB(textDescription) <> 0) Then cTextFile.AppendLineLf "TITLE """ & textDescription & """"
    
    'Tradition (?) dictates an extra blank line between some segments
    cTextFile.AppendLineBreakLf
    
    'LUT size gets its own line
    cTextFile.AppendLineLf "#LUT size"
    cTextFile.AppendLineLf "LUT_3D_SIZE " & Trim$(Str$(m_LutSize))
    cTextFile.AppendLineBreakLf
    
    'Domain needs to be explicitly written, per-channel
    cTextFile.AppendLineLf "#data domain"
    cTextFile.AppendLineLf "DOMAIN_MIN 0.0 0.0 0.0"
    cTextFile.AppendLineLf "DOMAIN_MAX 1.0 1.0 1.0"
    cTextFile.AppendLineBreakLf
    
    'We now append all colors in RGB order, with the R channel incrementing fastest.
    ' Colors are written as floating-point values written out to 6 decimal points.
    cTextFile.AppendLineLf "#LUT data points"
    
    Dim r As Long, g As Long, b As Long
    
    For b = 0 To m_LutLoopSize
    For g = 0 To m_LutLoopSize
    For r = 0 To m_LutLoopSize
        With m_lut3D(r, g, b)
            cTextFile.AppendLineLf Format$(.rFloat, "0.000000") & " " & Format$(.gFloat, "0.000000") & " " & Format$(.bFloat, "0.000000")
        End With
    Next r
    Next g
    Next b
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Constructing final cube text took " & VBHacks.GetTimeDiffNowAsString(startTime)
    
    'All that's left is to write the table out to file!
    If DEBUG_VERBOSE Then PDDebug.LogAction "Handing final cube text off to Files module for export to " & dstFile
    SaveLUTToFile_Cube = Files.FileSaveAsText(cTextFile.ToString, dstFile, True, False)
    
    Exit Function
    
CouldNotSaveFile:
    InternalProblem FUNC_NAME, "VB error #" & Err.Number & ": " & Err.Description
    SaveLUTToFile_Cube = False
    
End Function

'Export the current LUT to a look file
Friend Function SaveLUTToFile_look(ByRef dstFile As String, Optional ByRef copyrightText As String = vbNullString, Optional ByRef textDescription As String = vbNullString) As Boolean
    
    Const FUNC_NAME As String = "SaveLUTToFile_look"
    
    On Error GoTo CouldNotSaveFile
    
    'Failsafe checks for valid inputs and/or class state
    Dim cantSave As Boolean: cantSave = False
    
    If (LenB(dstFile) = 0) Then
        cantSave = True
        InternalProblem FUNC_NAME, "invalid function parameters"
    End If
    
    If (m_LutSize <= 0) Then
        cantSave = True
        InternalProblem FUNC_NAME, "no LUT"
    End If
    
    If cantSave Then
        SaveLUTToFile_look = False
        Exit Function
    End If
    
    'Failsafe check only; PD should enforce safe overwriting in the parent function.
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    Dim startTime As Currency
    If DEBUG_VERBOSE Then VBHacks.GetHighResTime startTime
    
    'To improve performance, write everything to a pdString instance.  Note that we'll be using
    ' Unix-style Lf line terminators (*not* CrLf) for compatibility reasons.  (Technically look files
    ' should work either way, but nearly all the ones I see "in the wild" have Unix line-endings.)
    Dim cTextFile As pdString
    Set cTextFile = New pdString
    
    'look files are basically XML files with the LUT itself embedded as hex.  Easy enough!
    ' (Note that look files coming from Adobe have a bunch of extra XML tags that we can't replicate,
    ' so we *only* write the LUT and a few key details.)
    cTextFile.AppendLineLf "<?xml version=""1.0"" ?>"
    cTextFile.AppendLineLf "<look>"
    
    'Try to mimic compatibility with Adobe where we can
    cTextFile.AppendLineLf "  <shaders>"
    cTextFile.AppendLineLf "  </shaders>"
    
    'Next comes the LUT block.  Start with an opening tag and size descriptor
    cTextFile.AppendLineLf Space$(2) & "<LUT>"
    cTextFile.AppendLineLf Space$(4) & "<size>""" & Trim$(Str$(m_LutSize)) & """</size>"
    cTextFile.AppendLineLf Space$(6) & "<data>"""
    
    'We now need to pull the full LUT matrix, as-is, into a hex string
    Dim lutHex As String
    If Not Strings.BytesToHex_FromPtr(VarPtr(m_lut3D(0, 0, 0)), (m_LutSize ^ 3) * 12, lutHex, True, True) Then
        InternalProblem FUNC_NAME, "bad hex conversion"
        SaveLUTToFile_look = False
        Exit Function
    End If
    
    Dim hexPrefix As String
    hexPrefix = Space$(6)
    
    'We now want to break up the hex string into 64-character chunks, and write each 64-char chunk as its own line
    Dim i As Long
    For i = 1 To Len(lutHex) Step 64
        
        Dim lenThisChunk As Long
        lenThisChunk = (Len(lutHex) - i) + 1
        If (lenThisChunk > 64) Then lenThisChunk = 64
        
        cTextFile.Append hexPrefix
        
        If (i < Len(lutHex) - 64) Then
            cTextFile.AppendLineLf Mid$(lutHex, i, lenThisChunk)
        Else
            cTextFile.AppendLineLf Mid$(lutHex, i, lenThisChunk) & """"
        End If
        
    Next i
    
    'Close all open tags, then write out the data!
    cTextFile.AppendLineLf Space$(4) & "</data>"
    cTextFile.AppendLineLf Space$(2) & "</LUT>"
    cTextFile.AppendLineLf "</look>"
    
    If DEBUG_VERBOSE Then
        PDDebug.LogAction "Constructing final look text took " & VBHacks.GetTimeDiffNowAsString(startTime)
        PDDebug.LogAction "Handing final look text off to Files module for export to " & dstFile
    End If
    
    SaveLUTToFile_look = Files.FileSaveAsText(cTextFile.ToString, dstFile, True, False)
    
    Exit Function
    
CouldNotSaveFile:
    InternalProblem FUNC_NAME, "VB error #" & Err.Number & ": " & Err.Description
    SaveLUTToFile_look = False
    
End Function

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
