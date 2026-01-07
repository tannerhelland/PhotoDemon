VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdCBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Comic Book Archive (CBZ) Container and Parser
'Copyright 2021-2026 by Tanner Helland
'Created: 18/June/21
'Last updated: 23/June/21
'Last update: wrap up initial build
'
'Comic book archives are simple .zip files containing a set of comic book pages in ascending order.
' (Note that non-zip variants also exist, like .CBR which use "RAR" format - PhotoDemon only supports
' CBZ files at present.)
'
'The format is loosely defined at Wikipedia:
' https://en.wikipedia.org/wiki/Comic_book_archive
'
'This class requires a copy of cZipArchive, an MIT-licensed zip library by wqweto@gmail.com.
' Many thanks to wqweto for not only sharing his class under a permissive license, but also being
' very responsive to bug reports and feature requests.  An original, un-altered copy of cZipArchive
' can be downloaded from GitHub, as can its attached MIT license (link good as of January 2021):
' https://github.com/wqweto/ZipArchive
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Some of our own internal parsers can perform perf timings;
' enable this value to send format-specific timing reports to PDDebug.
Private Const GENERATE_PERF_REPORTS As Boolean = False

'cZipArchive handles all zip interaction duties
Private WithEvents m_ZipArchive As cZipArchive
Attribute m_ZipArchive.VB_VarHelpID = -1

'When validating a CBZ file, we store the original filename that was validated; if the validation succeeds,
' we can reuse the m_ZipArchive instance for actual parsing duties.
Private m_Filename As String

Private m_OverrideFilename As String

'Validate a source filename as CBZ format.  Validation *does* touch the file - we look for a non-zero
' page count (PD doesn't support zero-page files).
Friend Function IsFileCBZ(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True) As Boolean
    
    If (m_ZipArchive Is Nothing) Then Set m_ZipArchive = New cZipArchive
    
    Dim potentiallyCBZ As Boolean
    potentiallyCBZ = True
    
    'Check extension up front, if requested
    If requireValidFileExtension Then potentiallyCBZ = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "CBZ", True)
    
    'Proceed with deeper validation as necessary
    If potentiallyCBZ Then
        
        'Attempt to load the archive
        potentiallyCBZ = m_ZipArchive.OpenArchive(srcFilename)
        If potentiallyCBZ Then
        
            'The file appears to be a valid zip archive.
            
            'Make sure at least one file is enclosed
            potentiallyCBZ = (m_ZipArchive.FileCount > 0)
            
            'On success, don't free the cZipArchive instance; we want to reuse the current instance
            ' for actual file import.  Instead, note the successfully validated filename so that the
            ' loader knows to skip validation steps.
            If potentiallyCBZ Then m_Filename = srcFilename
            
        End If
        
    End If
    
    IsFileCBZ = potentiallyCBZ
    
End Function

'Before loading an actual CBZ file, consider running IsFileCBZ(), above - this will validate the file for you,
' and you can avoid calling this function for non-CBZ files.
Friend Function LoadCBZ(ByRef srcFilename As String, ByRef dstImage As pdImage) As Boolean
    
    Const funcName As String = "LoadCBZ"
    
    LoadCBZ = False
    
    On Error GoTo CouldNotLoadFile
    
    'If we haven't validated the target file, do so now
    If Strings.StringsNotEqual(m_Filename, srcFilename) Then
        If (Not Me.IsFileCBZ(m_Filename)) Then
            InternalError "LoadCBZ", "target file isn't a comic book archive"
            Exit Function
        End If
    End If
    
    Dim startTime As Currency, firstTime As Currency
    Dim zipTime As Double, imgLoadTime As Double, postTime As Double
    If GENERATE_PERF_REPORTS Then
        VBHacks.GetHighResTime firstTime
        VBHacks.GetHighResTime startTime
    End If
    
    'The validator will have already loaded the target file, so this exists only as a failsafe.
    If (m_ZipArchive Is Nothing) Then
        Set m_ZipArchive = New cZipArchive
        m_ZipArchive.OpenArchive srcFilename
    End If
    
    If GENERATE_PERF_REPORTS Then
        zipTime = 0#
        imgLoadTime = 0#
        postTime = 0#
    End If
                    
    'Next, things are simple - start extracting pages!
    Dim numPagesAdded As Long
    numPagesAdded = 0
    
    'Note that pages can be in a variety of formats.  JPEG and PNG are the most common,
    ' but they're not guaranteed (e.g. I've seen GIF, BMP, and TIFF "in the wild").
    Dim cbPage As Long, pageData As Variant
    For cbPage = 0 To m_ZipArchive.FileCount - 1
        
        Message "Loading page %1 of %2...", CStr(cbPage + 1), m_ZipArchive.FileCount, "DONOTLOG"
        
        'Retrieve basic header information on the current file
        pageData = m_ZipArchive.FileInfo(cbPage)
        
        'Some CBZ writers place all pages in a parent folder; ignore this if it exists.
        ' (Note also that the weird Int cast and vbDirectory comparison is necessary to work around VB6
        '  quirks with bitwise operations and variants; an explicit cast ensures correct behavior.)
        Const idxAttributes As Long = 1
        If ((Int(pageData(idxAttributes)) And vbDirectory) <> vbDirectory) Then
        
            'This is not a folder.  It's most likely an image (page), but note that some users may
            ' acquire CBZ files from questionable sites.  Such archives may contain warez-style
            ' NFO files or other junk entries.  While I could check file extensions here,
            ' I'd rather err on the side of compatibility and simply extract the target file to a
            ' standalone temp file, *then* attempt to load it via PD's (extensive) standard
            ' file format coverage.
            
            'Start by generating a temp filename with the same extension as the target file.
            ' (Note that zip "filenames" can have illegal characters in filenames and other
            ' potential nonsense, so it's safer - and easier - to simply use a temp filename
            ' that we know is safe.)
            Dim tmpFilename As String
            tmpFilename = OS.UniqueTempFilename()
            
            Dim zipFilename As String
            Const idxFilename As Long = 0
            zipFilename = pageData(idxFilename)
            
            'Append the original file extension, if any, to the temp file
            If (LenB(Files.FileGetExtension(zipFilename)) <> 0) Then tmpFilename = tmpFilename & "." & Files.FileGetExtension(zipFilename)
            If GENERATE_PERF_REPORTS Then VBHacks.GetHighResTime startTime
            
            'Extract the target file.  (See the m_ZipArchive_BeforeExtract() event for the reason
            ' this m_OverrideFilename value exists.)
            m_OverrideFilename = tmpFilename
            If m_ZipArchive.Extract(tmpFilename, cbPage) Then
                
                If GENERATE_PERF_REPORTS Then
                    zipTime = zipTime + VBHacks.GetTimerDifferenceNow(startTime)
                    VBHacks.GetHighResTime startTime
                End If
                
                'Extraction was (probably) successful.  Attempt to load it into a temporary DIB.
                Dim imgOK As Boolean, tmpDIB As pdDIB
                imgOK = Loading.QuickLoadImageToDIB(tmpFilename, tmpDIB, False, False)
                    
                If GENERATE_PERF_REPORTS Then
                    imgLoadTime = imgLoadTime + VBHacks.GetTimerDifferenceNow(startTime)
                    VBHacks.GetHighResTime startTime
                End If
                
                'If the image is not in a format PD recognizes, that's fine - it's probably metadata
                ' or text info, which we can safely ignore.  (This function is currently set to
                ' *silently* ignore incompatible files, i.e. files that don't contain image data.)
                If imgOK Then
                
                    'Because color-management has already been handled (if applicable),
                    ' this is a great time to premultiply alpha.
                    If (Not tmpDIB.GetAlphaPremultiplication) Then tmpDIB.SetAlphaPremultiplication True
                    
                    'Prep a new layer object and initialize it with the image bits we've retrieved
                    Dim newLayerID As Long
                    newLayerID = dstImage.CreateBlankLayer()
                    
                    Dim tmpLayer As pdLayer
                    Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                    tmpLayer.InitializeNewLayer PDL_Image, Files.FileGetName(zipFilename), tmpDIB, True
                    
                    'Fill in any remaining layer properties
                    tmpLayer.SetLayerBlendMode BM_Normal
                    tmpLayer.SetLayerOpacity 100!
                    tmpLayer.SetLayerOffsetX 0
                    tmpLayer.SetLayerOffsetY 0
                    tmpLayer.SetLayerVisibility (numPagesAdded = 0)
                    numPagesAdded = numPagesAdded + 1
                    
                    'Notify the layer and image of new changes, so it knows to regenerate internal caches
                    ' on next access
                    tmpLayer.NotifyOfDestructiveChanges
                    dstImage.NotifyImageChanged UNDO_Image
                    
                    If GENERATE_PERF_REPORTS Then postTime = postTime + VBHacks.GetTimerDifferenceNow(startTime)
                    
                '/end imgOK
                Else
                    InternalError funcName, "skipped page """ & pageData(idxFilename) & """ (not an image file)"
                End If
            
            '/end successful extraction
            Else
                InternalError funcName, "couldn't extract """ & pageData(idxFilename) & """ to """ & tmpFilename & """"
            End If
            
        '/end Not isDirectory
        Else
            'If you want directory information, you could retrieve it here:
            'PDDebug.LogAction "folder: " & pageData(idxFilename)
        End If
        
    Next cbPage

    'All layers have been iterated.  If the target image contains at least one valid layer,
    ' consider this a successful load.
    LoadCBZ = (dstImage.GetNumOfLayers > 0)
    If LoadCBZ Then
    
        'Activate the base layer (typically the first page) and make it the only visible page
        dstImage.SetActiveLayerByIndex 0
        
        'Fit the canvas around the largest layer (important if two-page spreads are present)
        CalculateImageSize dstImage
        
    End If
    
    If GENERATE_PERF_REPORTS Then
        InternalWarning "LoadCBZ", "Total time to load CBZ file: " & VBHacks.GetTimeDiffNowAsString(firstTime)
        InternalWarning "LoadCBZ", "Time spent in zip extraction: " & Format$(zipTime * 1000#, "0.0") & " ms"
        InternalWarning "LoadCBZ", "Time spent in image file parsing: " & Format$(imgLoadTime * 1000#, "0.0") & " ms"
        InternalWarning "LoadCBZ", "Time spent in post-processing: " & Format$(postTime * 1000#, "0.0") & " ms"
    End If
    
    Set m_ZipArchive = Nothing
    
    Exit Function
    
CouldNotLoadFile:
    InternalError "LoadCBZ", "Internal VB error #" & Err.Number & ": " & Err.Description
    LoadCBZ = False

End Function

Private Sub CalculateImageSize(ByRef dstImage As pdImage)
        
    'Start by finding two things:
    ' 1) The lowest x/y offsets in the current layer stack
    ' 2) The highest width/height in the current layer stack (while accounting for offsets as well!)
    Dim maxWidth As Long, maxHeight As Long
    maxWidth = 0: maxHeight = 0
    
    Dim i As Long
    For i = 0 To dstImage.GetNumOfLayers - 1
        
        'Retrieve layer dimensions, and compare against current max sizes
        If (dstImage.GetLayerByIndex(i).GetLayerWidth() > maxWidth) Then maxWidth = dstImage.GetLayerByIndex(i).GetLayerWidth()
        If (dstImage.GetLayerByIndex(i).GetLayerHeight() > maxHeight) Then maxHeight = dstImage.GetLayerByIndex(i).GetLayerHeight()
        
    Next i
    
    'Now that we know a maximum size, assign it to the parent image, then center all pages within
    ' the new image boundaries.
    dstImage.SetDPI 96#, 96#
    dstImage.Width = maxWidth
    dstImage.Height = maxHeight
    
    For i = 0 To dstImage.GetNumOfLayers - 1
        With dstImage.GetLayerByIndex(i)
            .SetLayerOffsetX (maxWidth - .GetLayerWidth()) \ 2
            .SetLayerOffsetY (maxHeight - .GetLayerHeight()) \ 2
        End With
    Next i
    
    dstImage.NotifyImageChanged UNDO_Image
    
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdCBZ." & funcName & "() reported an error on file """ & m_Filename & """: " & errDescription
    Else
        Debug.Print "pdCBZ." & funcName & "() reported an error on file """ & m_Filename & """: " & errDescription
    End If
End Sub

Private Sub InternalWarning(ByRef funcName As String, ByRef warnText As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdCBZ." & funcName & "() warned: " & warnText
    Else
        Debug.Print "pdCBZ." & funcName & "() warned: " & warnText
    End If
End Sub

'cZipArchive tries to be helpful by extracting files using their original filenames, appended to whatever path
' you pass to them.  For privacy reasons, we want to use random filenames
Private Sub m_ZipArchive_BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean, Cancel As Boolean)
    File = m_OverrideFilename
End Sub
