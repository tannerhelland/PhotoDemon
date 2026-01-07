Attribute VB_Name = "ImageFormats_GIF"
'***************************************************************************
'Additional support functions for GIF support
'Copyright 2001-2026 by Tanner Helland
'Created: 4/15/01
'Last updated: 21/June/22
'Last update: unify debug text across the codebase
'Dependencies: pdGIF (class), ImageFormats_GIF_LZW (module)
'
'Most image exporters exist in the ImageExporter module.  GIF is a weird exception because animated GIFs
' require a ton of preprocessing (to optimize animation frames), so I've moved them to their own home.
'
'PhotoDemon automatically optimizes saved GIFs to produce the smallest possible files.  A variety of
' optimizations are used, and the encoder tests various strategies to try and choose the "best"
' (smallest) solution on each frame.  As you can see from the size of this module, many many many
' different optimizations are attempted.  Despite this, the optimization pre-pass is reasonably quick,
' and the GIFs produced this way are often an order of magnitude (or more) smaller than a naive
' GIF encoder would produce.
'
'Note that the optimization steps are specifically written in an export-library-agnostic way.
' PD internally stores the results of all optimizations, then just hands the optimized frames off
' to an encoder at the end of the process.  Historically PD used FreeImage for animated GIF encoding,
' but FreeImage has a number of shortcomings (including woeful performance and writing larger GIFs
' than is necessary), so in 2021 we moved to an in-house LZW encoder based off the classic UNIX
' "compress" tool.  The LZW encoder lives in a separate module (ImageFormats_GIF_LZW).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Low-level static GIF export interface.  As of 2021, image pre-processing (including palettization)
' and GIF encoding is all performed using homebrew code.
Public Function ExportGIF_LL(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportGIFError
    
    ExportGIF_LL = False
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to
    ' a new file, and if it's saved successfully, overwrite the original file - this way,
    ' if an error occurs mid-save, the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    'As always, pdStream handles actual writing duties.  (Memory mapping is used for ideal performance.)
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, tmpFilename, optimizeAccess:=OptimizeSequentialAccess) Then
        
        'A pdGIF instance handles the actual encoding
        Dim cGIF As pdGIF
        Set cGIF = New pdGIF
        If cGIF.SaveGIF_ToStream_Static(srcPDImage, cStream, formatParams, metadataParams) Then
            
            'Close the stream, then release the pdGIF instance
            cStream.StopStream
            Set cGIF = Nothing
            
            'If we wrote our data to a temp file, attempt to replace the original file
            If Strings.StringsNotEqual(dstFile, tmpFilename) Then
                
                ExportGIF_LL = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
                
                If (Not ExportGIF_LL) Then
                    Files.FileDelete tmpFilename
                    PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                End If
            
            'Encode is already done!
            Else
                ExportGIF_LL = True
            End If
            
        Else
            PDDebug.LogAction "WARNING! pdGIF failed to save GIF"
        End If
        
        ProgressBars.SetProgBarVal 0
        ProgressBars.ReleaseProgressBar
        
    Else
        PDDebug.LogAction "WARNING!  Couldn't initialize stream against " & dstFile
    End If
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error #" & Err.Number & ", " & Err.Description
    ExportGIF_LL = False
    
End Function

'Low-level animated GIF export interface.  As of 2021, image pre-processing (including palettization)
' and GIF encoding is all performed using homebrew code.
Public Function ExportGIF_Animated_LL(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportGIFError
    
    ExportGIF_Animated_LL = False
     
    'Initialize a progress bar.  Animated GIFs are saved in two (technically three) passes.
    ' 1) The first pass handles most the pre-processing work.  It attempts to assemble a global (shared)
    '    palette for the image, which can be used by multiple frames.  If a frame's color requirements
    '    exceed the size limit of the global palette, the frame will get a local palette instead.
    '    Non-palette-related duties like frame cropping are also handled in this step, and frames with
    '    a local palette are fully optimized (because we know everything we need to know to optimize).
    ' 2) The second pass handles remaining optimization duties for frames that use the (now complete)
    '    global palette.  Because the full contents of the global palette may not be known until all
    '    frames are analyzed, we couldn't do optimizations like pixel-blanking on these frames yet.
    '    At the end of this pass, all frames are fully assembled, all frame properties are known,
    '    and the frames are ready to be dumped to file.
    ' 3) The final pass is the actual writing of the GIF file.  This pass is so fast that we don't
    '    bother using the progress bar for it.  (Even on an animation with hundreds of frames and a
    '    large total image size, this step typically occurs in <= 1 second.)
    ProgressBars.SetProgBarMax srcPDImage.GetNumOfLayers * 2
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to
    ' a new file, and if it's saved successfully, overwrite the original file - this way,
    ' if an error occurs mid-save, the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    'As always, pdStream handles actual writing duties.  (Memory mapping is used for ideal performance.)
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, tmpFilename, optimizeAccess:=OptimizeSequentialAccess) Then
        
        'A pdGIF instance handles the actual encoding
        Dim cGIF As pdGIF
        Set cGIF = New pdGIF
        If cGIF.SaveGIF_ToStream_Animated(srcPDImage, cStream, formatParams, metadataParams) Then
            
            'Close the stream, then release the pdGIF instance
            cStream.StopStream
            Set cGIF = Nothing
            
            'If we wrote our data to a temp file, attempt to replace the original file
            If Strings.StringsNotEqual(dstFile, tmpFilename) Then
                
                ExportGIF_Animated_LL = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
                
                If (Not ExportGIF_Animated_LL) Then
                    Files.FileDelete tmpFilename
                    PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                End If
            
            'Encode is already done!
            Else
                ExportGIF_Animated_LL = True
            End If
            
        Else
            PDDebug.LogAction "WARNING! pdGIF failed to save GIF"
        End If
        
        ProgressBars.SetProgBarVal 0
        ProgressBars.ReleaseProgressBar
        
    Else
        PDDebug.LogAction "WARNING!  Couldn't initialize stream against " & dstFile
    End If
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error #" & Err.Number & ", " & Err.Description
    ExportGIF_Animated_LL = False
    
End Function

Private Sub ExportDebugMsg(ByRef srcDebugMsg As String)
    PDDebug.LogAction srcDebugMsg
End Sub
