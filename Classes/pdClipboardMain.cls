VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdClipboardMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Clipboard Interface
'Copyright 2001-2026 by Tanner Helland
'Created: 15/April/01
'Last updated: 08/April/25
'Last update: fix behavior when copy+pasting internal data formats (like layers with active affine transforms)
'
'Module for handling all Windows clipboard routines.  Most clipboard APIs are not located here,
' but in the separate pdClipboard object, which includes a ton of specialized helper functions.
' Look there for details.
'
'Special thanks go out to LaVolpe for his very helpful work in coercing VB into accepting
' Unicode-friendly drag/drop targets. In particular, these two references were invaluable:
' http://www.vbforums.com/showthread.php?637335-RESOLVED-Drag-amp-Drop-file-from-compressed-folder&p=3943417&viewfull=1#post3943417
' http://cyberactivex.com/UnicodeTutorialVb.htm#Filenames_via_DragDrop_or_Paste
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'**************************************************************************
'CONSTANTS FOR DEBUGGING USE ONLY.
' In production code, these should all be set to TRUE.
Private Const PD_CB_ALLOW_BMP_DIB_PASTE As Boolean = True
Private Const PD_CB_ALLOW_FILE_PASTE As Boolean = True
Private Const PD_CB_ALLOW_GENERIC_TEXT_PASTE As Boolean = True
Private Const PD_CB_ALLOW_INTERNAL_FORMAT_PASTE As Boolean = True
Private Const PD_CB_ALLOW_JPEG_PASTE As Boolean = True
Private Const PD_CB_ALLOW_METAFILE_PASTE As Boolean = True
Private Const PD_CB_ALLOW_ONLINE_PASTE As Boolean = True
Private Const PD_CB_ALLOW_PNG_PASTE As Boolean = True
'**************************************************************************

'Image formats that PD understands (via clipboard)
Public Enum PD_ClipboardFormats
    pdcf_All = 0
    pdcf_InternalPD = CF_PD_DIB
    pdcf_InternalLayer = CF_PD_LAYER
    pdcf_Bitmap = CF_BITMAP
    pdcf_Dib = CF_DIB
    pdcf_DibV5 = CF_DIBV5
    pdcf_PNG = -1
End Enum

#If False Then
    Private Const pdcf_All = 0, pdcf_InternalPD = CF_PD_DIB, pdcf_InternalLayer = CF_PD_LAYER, pdcf_Bitmap = CF_BITMAP, pdcf_Dib = CF_DIB, pdcf_DibV5 = CF_DIBV5, pdcf_PNG = -1
#End If

'Some format constants are not inherently defined by VB; we cache these at load-time
Private CF_FileContents As Integer, CF_FileGroupDescriptorW As Integer

'Some specialized structs are required for parsing various clipboard bitmap formats
Private Type BITMAPFILEHEADER
    Type As Integer
    Size As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffBits As Long
End Type

Private Enum BMP_COMPRESSION
    BC_RGB = 0
    BC_RLE8 = 1
    BC_RLE4 = 2
    BC_BITFIELDS = 3
    BC_JPEG = 4
    BC_PNG = 5
End Enum

#If False Then
    Private Const BC_RGB = 0, BC_RLE8 = 1, BC_RLE4 = 2, BC_BITFIELDS = 3, BC_JPEG = 4, BC_PNG = 5
#End If

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As BMP_COMPRESSION    'Note that this enum only works for still frames; video has different values, a la https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPV5HEADER         ' Offset from start of struct
    biSize As Long                  ' 0
    biWidth As Long                 ' 4
    biHeight As Long                ' 8
    biPlanes As Integer             ' 12
    biBitCount As Integer           ' 14
    biCompression As Long           ' 16
    biSizeImage As Long             ' 20
    biXPelsPerMeter As Long         ' 24
    biYPelsPerMeter As Long         ' 28
    biClrUsed As Long               ' 32
    biClrImportant As Long          ' 36 (NOTE: Default BitmapInfoHeader struct ends here)
    biRedMask As Long               ' 40
    biGreenMask As Long             ' 44
    biBlueMask As Long              ' 48
    biAlphaMask As Long             ' 52
    biCSType As Long                ' 56
    CIEXYZ_RX As Long               ' 60 (NOTE: CIEXYZTRIPLE structures exist for each of R, G, B.  XYZ triples use a bizarre custom format,
    CIEXYZ_RY As Long               ' 64         so you can't actually use these Long values as-is; they need further processing!)
    CIEXYZ_RZ As Long               ' 68
    CIEXYZ_GX As Long               ' 72
    CIEXYZ_GY As Long               ' 76
    CIEXYZ_GZ As Long               ' 80
    CIEXYZ_BX As Long               ' 84
    CIEXYZ_BY As Long               ' 88
    CIEXYZ_BZ As Long               ' 92
    biGammaRed As Long              ' 96
    biGammaGreen As Long            ' 100
    biGammaBlue As Long             ' 104 (NOTE: BitmapV4Header struct ends here)
    biIntent As Long                ' 108
    biProfileData As Long           ' 112
    biProfileSize As Long           ' 116
    biReserved As Long              ' 120 (NOTE: BitmapV5Header struct ends here)
End Type

Private Const RGBA_RED_MASK As Long = &HFF0000
Private Const RGBA_GREEN_MASK As Long = &HFF00&
Private Const RGBA_BLUE_MASK As Long = &HFF&
Private Const RGBA_ALPHA_MASK As Long = &HFF000000

Private Declare Function CopyEnhMetaFileW Lib "gdi32" (ByVal hEmfSrc As Long, ByVal ptrToDstFilename As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

'If the clipboard is currently open, this module will cache certain values.  External functions (like the image load routines) can query these
' for additional information on how to handle the image.  Clipboard images have a ton of caveats that normal images do not; hence these trackers.
Private m_IsClipboardOpen As Boolean

'If external functions want to know more about the current clipboard operation, they can retrieve a copy of this struct.
Private m_ClipboardInfo As PD_Clipboard_Info

'Lots of clipboard actions require a temporary DIB
Private m_ClipboardDIB As pdDIB

'Stashing layer data will generate a layer header and vector stash as well
Private m_StashedLayer As String, m_StashedVector As String

'When stashing clipboard data, we use a temp file.
' Note that certain internal types of copy+paste (e.g. copy+pasting a vector layer, or a layer with active affine transforms)
' may stash multiple types of data.
Private m_StashFile As String, m_StashFileLayer As String, m_StashFileLayerAffineOrig As String, m_StashFileVector As String

'Clipboard interaction object; importantly, if delayed rendering is active, this class *will* raise events that must be handled.
Private WithEvents m_Clipboard As pdClipboard
Attribute m_Clipboard.VB_VarHelpID = -1

'Next comes all the helper functions for Unicode-enabled drag/drop behavior.
' (Thank you to LaVolpe for these functions, c/o http://cyberactivex.com/UnicodeTutorialVb.htm#Filenames_via_DragDrop_or_Paste
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As VbVarType, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal ptrToFirstChar As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub ReleaseStgMedium Lib "ole32" (ByVal ptrToStgMedium As Long)

Private Const TYMED_HGLOBAL As Long = 1
Private Const TYMED_ISTREAM As Long = 4
Private Const DVASPECT_CONTENT As Long = 1
Private Const IDataObjVTable_GetData As Long = 12   '12 is an offset to the GetData function (e.g. the 4th VTable entry)
Private Const CC_STDCALL As Long = 4

Private Type FORMATETC
    cfFormat As Long
    pDVTARGETDEVICE As Long
    dwAspect As Long
    lIndex As Long
    TYMED As Long
End Type

Private Type STGMEDIUM
    TYMED As Long
    Data As Long
    pUnkForRelease As Long
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SIZEL
    cx As Long
    cy As Long
End Type
Private Type POINTL
    x As Long
    y As Long
End Type

Private Type FILEDESCRIPTOR
    dwFlags As Long
    clsID(0 To 3) As Long
    cSize As SIZEL
    cPoint As POINTL
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    cFileName(0 To 519) As Byte   'MAXPATH wchars
End Type

'When pasting new layers in a given session, we increment a "paste" counter that we can use
' to modify layer names - e.g. "clipboard layer 8" for the 8th pasted layer in a session.
Private m_LayerCounter As Long

Friend Function IsClipboardOpen() As Boolean
    IsClipboardOpen = m_IsClipboardOpen
End Function

Friend Function GetClipboardInfo() As PD_Clipboard_Info
    GetClipboardInfo = m_ClipboardInfo
End Function

'Copy the current selection (or entire layer, if no selection is active) to the clipboard, then erase the selected area
' (or entire layer, if no selection is active).
Friend Sub ClipboardCut(ByVal cutMerged As Boolean, Optional ByVal cFormat As PD_ClipboardFormats = pdcf_All)
    
    'Lock the UI
    Processor.MarkProgramBusyState True, True
    
    'Cut begins as a normal Copy operation
    ClipboardCopy cutMerged, False, cFormat
    
    'Once the copy is complete, we take the extra step of erasing the selected area from the screen.  Note that "Cut merged" requires us
    ' to delete the selected region from *all visible layers*, so in advance, let's figure out which layers are affected.
    Dim startLayer As Long, endLayer As Long
    
    If cutMerged Then
        startLayer = 0
        endLayer = PDImages.GetActiveImage.GetNumOfLayers - 1
    Else
        startLayer = PDImages.GetActiveImage.GetActiveLayerIndex
        endLayer = PDImages.GetActiveImage.GetActiveLayerIndex
    End If
    
    Dim i As Long
    For i = startLayer To endLayer
        
        'For "cut merged", ignore transparent layers
        If cutMerged Then
        
            If PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility Then
                
                'If a selection is active, erase the selected area.  Otherwise, wipe the whole layer.
                If PDImages.GetActiveImage.IsSelectionActive Then
                    PDImages.GetActiveImage.EraseProcessedSelection i
                Else
                    Layers.EraseLayerByIndex i
                End If
                
            End If
        
        'For "cut from layer", erase the selection regardless of layer visibility
        Else
        
            'If a selection is active, erase the selected area.  Otherwise, delete the given layer.
            If PDImages.GetActiveImage.IsSelectionActive Then
                PDImages.GetActiveImage.EraseProcessedSelection i
            Else
                Layers.DeleteLayer i
            End If
            
        End If
        
    Next i
    
    'Unlock the UI
    Processor.MarkProgramBusyState False
    
    'Redraw the active viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

Friend Sub ClipboardCutSpecial(ByRef paramString As String)

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString paramString
    
    With cParams
        Me.ClipboardCut cParams.GetBool("merged", True), cParams.GetLong("clipboard-format", pdcf_All)
    End With

End Sub

'Copy the current layer (or composite image, if copyMerged is true) to the clipboard.
' The optional updateUI parameter tells the function to lock PD's UI during the copy operation.  ("Cut" operations also do this, so they
' manage the UI separately - hence the need for a dedicated parameter.)
Friend Sub ClipboardCopy(ByVal copyMerged As Boolean, Optional ByVal updateUI As Boolean = True, Optional ByVal cFormat As PD_ClipboardFormats = pdcf_All)
    
    'Lock the UI
    If updateUI Then Processor.MarkProgramBusyState True, True
    
    'Stash the relevant clipboard data to file
    StashClipboardData copyMerged
    
    'Notify the clipboard of the formats we are willing to copy "on-demand"
    If m_Clipboard.ClipboardOpen(FormMain.hWnd) Then
        
        m_Clipboard.ClearClipboard
        
        'By default, we offer all supported formats
        If (cFormat = pdcf_All) Then
        
            'Internal pixel data is always available
            m_Clipboard.SetClipboardData_DelayedRendering CF_PD_DIB
            
            'If the copy is of a whole layer (not a selection or merged image), we'll also save out a full
            ' layer header and any relevant vector data so that we can re-create the whole layer as-is.
            If (Not copyMerged) And (Not PDImages.GetActiveImage.IsSelectionActive) Then m_Clipboard.SetClipboardData_DelayedRendering CF_PD_LAYER
            
            'Bitmap formats are for other apps
            m_Clipboard.SetClipboardData_DelayedRendering m_Clipboard.AddClipboardFormat("PNG")
            m_Clipboard.SetClipboardData_DelayedRendering CF_BITMAP
            m_Clipboard.SetClipboardData_DelayedRendering CF_DIBV5
            
            'So an interesting story about CF_DIB.  PD is perfectly capable of rendering CF_DIB images
            ' to the clipboard, but at present, we choose not to enable this.  Why?  Microsoft Paint will
            ' preferentially accept a CF_DIB image over any other format on the clipboard.  This would be
            ' okay if Paint recognized alpha channels, but because it doesn't, we would need to downsample
            ' PD's 32-bpp images to 24-bpp.  However, this would incorrectly generate a 24-bpp copy for
            ' software that *does* recognize 32-bpp CF_DIB entries.  To minimize the risk of a screwed-up
            ' end result, PD only registers CF_BITMAP and CF_DIBv5 formats to the clipboard (plus PNG,
            ' obviously).  It's assumed that any program clever enough to recognize CF_DIBv5 also has a
            ' contingency plan for alpha bytes, whereas the same can't be said for CF_DIB, unfortunately.
            '
            'As always, PNG is the preferred interchange format for images.  If you use a program that
            ' doesn't copy/paste PNG-format data, get them to fix their software!
            
            'UPDATE December 2020: due to clipboard issues in Google Chrome, I've reenabled explicit DIB
            ' availability on the clipboard.  This allows pasting into Chrome, and my hope is that as of
            ' this year, any programs that support alpha channels will be smart enough to grab PNG or
            ' DIBv5 data instead of bare DIB data (which is pre-composited against a white background to
            ' ensure correct behavior in alpha-unaware software like MS Paint).  This does break 32-bpp
            ' pasting into some legacy software.  As a workaround, PD's Edit > Special Copy/Cut menu allows
            ' you to explicitly paste DIBv5 data which is an expensive workaround, but hey - at least it's
            ' there if you need it.
            m_Clipboard.SetClipboardData_DelayedRendering CF_DIB
            
        ElseIf (cFormat = pdcf_Bitmap) Then
            m_Clipboard.SetClipboardData_DelayedRendering CF_BITMAP
        ElseIf (cFormat = pdcf_Dib) Then
            m_Clipboard.SetClipboardData_DelayedRendering CF_DIB
        ElseIf (cFormat = pdcf_DibV5) Then
            m_Clipboard.SetClipboardData_DelayedRendering CF_DIBV5
        ElseIf (cFormat = pdcf_InternalPD) Then
            m_Clipboard.SetClipboardData_DelayedRendering CF_PD_DIB
        ElseIf (cFormat = pdcf_InternalLayer) Then
            m_Clipboard.SetClipboardData_DelayedRendering CF_PD_LAYER
        ElseIf (cFormat = pdcf_PNG) Then
            m_Clipboard.SetClipboardData_DelayedRendering m_Clipboard.AddClipboardFormat("PNG")
        Else
            PDDebug.LogAction "WARNING!  pdClipboardMain.ClipboardCopy received an unsupported format request."
        End If
        
        m_Clipboard.ClipboardClose
        
    Else
        PDDebug.LogAction "WARNING!  pdClipboardMain.ClipboardCopy couldn't open the clipboard!"
    End If
    
    'Unlock the UI
    If updateUI Then Processor.MarkProgramBusyState False
    
End Sub

Friend Sub ClipboardCopySpecial(ByRef paramString As String)

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString paramString
    
    With cParams
        Me.ClipboardCopy cParams.GetBool("merged", True), True, cParams.GetLong("clipboard-format", pdcf_All)
    End With
    
End Sub

'Sometimes (mostly during debugging) it's helpful to copy arbitrary text data to the clipboard.
' This function will do that, taking care to preserve Unicode text.
Friend Sub ClipboardCopy_Text(ByRef srcString As String)
    If m_Clipboard.ClipboardOpen(FormMain.hWnd) Then
        m_Clipboard.ClearClipboard
        m_Clipboard.SetClipboardBinaryData_Ptr CF_UNICODETEXT, StrPtr(srcString), LenB(srcString) + 2
        m_Clipboard.ClipboardClose
    End If
End Sub

Friend Sub ClipboardEmpty()
    If m_Clipboard.ClipboardOpen(FormMain.hWnd) Then
        m_Clipboard.ClearClipboard
        m_Clipboard.ClipboardClose
    End If
End Sub

'Paste an image (e.g. create new image data based on whatever is in the clipboard).
' The parameter "srcIsMeantAsLayer" controls whether the clipboard data is loaded as a new image, or as a new layer in an existing image.
Friend Function ClipboardPaste(ByVal srcIsMeantAsLayer As Boolean, Optional ByRef pasteToThisDIBInstead As pdDIB = Nothing, Optional ByRef fullParamString As String = vbNullString) As Boolean
    
    Dim pasteWasSuccessful As Boolean
    pasteWasSuccessful = False
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString fullParamString
    
    'Attempt to open the clipboard
    If m_Clipboard.ClipboardOpen(FormMain.hWnd) Then
        
        'Mark the clipboard as open; external functions can query this value
        m_IsClipboardOpen = True
        
        'When debugging, it's nice to know what clipboard formats the OS reports prior to actually retrieving them.
        PDDebug.LogAction "Clipboard reports the following formats: " & m_Clipboard.GetListOfAvailableFormatNames()
        
        'If PD was used to Cut or Copy something onto the clipboard, our own private format(s) will be listed first.
        If (m_Clipboard.DoesClipboardHaveFormatID(CF_PD_DIB) Or m_Clipboard.DoesClipboardHaveFormatID(CF_PD_LAYER)) And PD_CB_ALLOW_INTERNAL_FORMAT_PASTE Then
            pasteWasSuccessful = ClipboardPaste_InternalData(srcIsMeantAsLayer, pasteToThisDIBInstead)
        End If
        
        'PNGs on the clipboard get preferential treatment, as they preserve transparency data - so check for them first.
        If m_Clipboard.DoesClipboardHaveFormatName("PNG") And (Not pasteWasSuccessful) And PD_CB_ALLOW_PNG_PASTE Then
            pasteWasSuccessful = ClipboardPaste_CustomImageFormat("PNG", srcIsMeantAsLayer, "png", pasteToThisDIBInstead)
        End If
        
        'If we couldn't find PNG data (or something went horribly wrong during that step), look for an HTML fragment next.
        ' Images copied from web browsers typically create an HTML fragment, which should have a direct link to the copied image.
        '  Downloading the image manually lets us maintain things like ICC profiles and the image's original filename.
        If m_Clipboard.DoesClipboardHaveHTML() And (Not pasteWasSuccessful) And PD_CB_ALLOW_ONLINE_PASTE Then
            pasteWasSuccessful = ClipboardPaste_HTML(srcIsMeantAsLayer, pasteToThisDIBInstead)
        End If
        
        'JPEGs are another possibility.  We prefer them less than PNG or direct download (because there's no guarantee that the
        ' damn browser didn't re-encode them, but they're better than bitmaps or DIBs because they may retain metadata and
        ' color profiles, so test for JPEG next.  (Also, note that certain versions of Microsoft Office use "JFIF" as the identifier,
        ' for reasons known only to them...)
        If m_Clipboard.DoesClipboardHaveFormatName("JPEG") And (Not pasteWasSuccessful) And PD_CB_ALLOW_JPEG_PASTE Then
            pasteWasSuccessful = ClipboardPaste_CustomImageFormat("JPEG", srcIsMeantAsLayer, "jpg", pasteToThisDIBInstead)
        End If
        
        If m_Clipboard.DoesClipboardHaveFormatName("JPG") And (Not pasteWasSuccessful) And PD_CB_ALLOW_JPEG_PASTE Then
            pasteWasSuccessful = ClipboardPaste_CustomImageFormat("JPG", srcIsMeantAsLayer, "jpg", pasteToThisDIBInstead)
        End If
        
        If m_Clipboard.DoesClipboardHaveFormatName("JFIF") And (Not pasteWasSuccessful) And PD_CB_ALLOW_JPEG_PASTE Then
            pasteWasSuccessful = ClipboardPaste_CustomImageFormat("JFIF", srcIsMeantAsLayer, "jpg", pasteToThisDIBInstead)
        End If
        
        'Next, see if the clipboard contains a generic file list.  If it does, try to load each file in turn.
        If m_Clipboard.DoesClipboardHaveFiles() And (Not pasteWasSuccessful) And PD_CB_ALLOW_FILE_PASTE Then
            pasteWasSuccessful = ClipboardPaste_ListOfFiles(srcIsMeantAsLayer, pasteToThisDIBInstead)
        End If
        
        'Next, look for plaintext.  This could be a URL, or maybe a text representation of a filepath.
        ' (Also, note that we only have to search for one text format, because the OS auto-converts between text formats for free.)
        If m_Clipboard.DoesClipboardHaveText() And (Not pasteWasSuccessful) And PD_CB_ALLOW_GENERIC_TEXT_PASTE Then
            pasteWasSuccessful = ClipboardPaste_TextSource(srcIsMeantAsLayer, pasteToThisDIBInstead)
        End If
        
        'Next, use any DIBs or bitmaps.  Once again, the OS auto-converts between bitmap and DIB formats, and if it all possible,
        ' we prefer DIBv5 as it actually supports alpha data.
        If m_Clipboard.DoesClipboardHaveBitmapImage() And (Not pasteWasSuccessful) And PD_CB_ALLOW_BMP_DIB_PASTE Then
            pasteWasSuccessful = ClipboardPaste_BitmapImage(srcIsMeantAsLayer, pasteToThisDIBInstead)
        End If
        
        'As a last resort, try to grab metafiles.  These are not ideal as some software (*cough* OFFICE *cough*) generates
        ' pretty terrible metafiles, but it's better than nothing.
        If m_Clipboard.DoesClipboardHaveMetafile() And (Not pasteWasSuccessful) And PD_CB_ALLOW_METAFILE_PASTE Then
            pasteWasSuccessful = ClipboardPaste_Metafile(srcIsMeantAsLayer, pasteToThisDIBInstead)
        End If
        
        'Regardless of success or failure, make sure to close the clipboard now that we're done with it.
        m_Clipboard.ClipboardClose
        
        'Mark the clipboard as closed
        m_IsClipboardOpen = False
        
    Else
        PDDebug.LogAction "WARNING!  Couldn't open the clipboard; is it possible another program has locked it?"
    End If
    
    ClipboardPaste = pasteWasSuccessful
    
    'If a paste operation was successful, switch the current tool to the layer move/resize tool,
    ' which is most likely needed after a new layer has been pasted.
    If pasteWasSuccessful Then
        
        If srcIsMeantAsLayer And (pasteToThisDIBInstead Is Nothing) Then
            
            'Check the param string; the Paste to Cursor command specifies a manual canvas x/y to paste to
            Dim positionedAlready As Boolean
            positionedAlready = False
            
            'If passed mouse positions exist, they will be in hWnd coordinate space, and they will only exist if
            ' the mouse was over the canvas at the time of loading (hypothetically - testing is TODO).
            If cParams.DoesParamExist("canvas-mouse-x", True) And cParams.DoesParamExist("canvas-mouse-y") Then
            
                'Mouse coordinates exist!  Retrieve and validate them, then apply them to
                ' the newly created layer.
                Dim reportedX As Long, reportedY As Long
                reportedX = cParams.GetLong("canvas-mouse-x", PDImages.GetActiveImage.GetActiveLayer.GetLayerOffsetX(), True)
                reportedY = cParams.GetLong("canvas-mouse-y", PDImages.GetActiveImage.GetActiveLayer.GetLayerOffsetY(), True)
                
                'Convert those values to image coordinates
                Dim imgX As Double, imgY As Double
                Drawing.ConvertCanvasCoordsToImageCoords FormMain.MainCanvas(0), PDImages.GetActiveImage, reportedX, reportedY, imgX, imgY, True
                
                'Apply to the layer
                PDImages.GetActiveImage.GetActiveLayer.SetLayerOffsetX Int(imgX + 0.5)
                PDImages.GetActiveImage.GetActiveLayer.SetLayerOffsetY Int(imgY + 0.5)
                
                'Notify the parent image of the change
                PDImages.GetActiveImage.NotifyImageChanged UNDO_LayerHeader, PDImages.GetActiveImage.GetActiveLayerIndex
                positionedAlready = True
                
            End If
            
            'If we didn't position the layer in a previous step, the newly pasted layer will default
            ' to the top-left corner of the current viewport (or (0, 0), if the viewport's corner is
            ' outside image boundaries).  To increase the chance of positioning the new layer in a
            ' "useful" location, look for an active selection.  If we find one, position this layer
            ' to the selection's top-left corner - this makes copy+paste operations produce
            ' identically position segments of the image, which is likely more useful than just
            ' dumping the newly pasted layer elsewhere.
            If (Not positionedAlready) And PDImages.GetActiveImage.IsSelectionActive Then
                
                If PDImages.GetActiveImage.MainSelection.IsLockedIn() Then
                
                    Dim selRect As RectF
                    selRect = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect()
                    
                    'Apply to the layer
                    PDImages.GetActiveImage.GetActiveLayer.SetLayerOffsetX Int(selRect.Left + 0.5)
                    PDImages.GetActiveImage.GetActiveLayer.SetLayerOffsetY Int(selRect.Top + 0.5)
                    
                    'Notify the parent image of the change
                    PDImages.GetActiveImage.NotifyImageChanged UNDO_LayerHeader, PDImages.GetActiveImage.GetActiveLayerIndex
                    positionedAlready = True
                    
                End If
                
            End If
            
            'Render the new image to screen
            Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
            
            'Synchronize the interface to the new image
            SyncInterfaceToCurrentImage
            
            'Finally, activate the move tool (in case the user wants to further tweak position).
            ' (NOTE Nov 2025: this behavior is now user-controlled by a Tools > Options > Interface preference)
            If UserPrefs.GetPref_Boolean("Interface", "MoveToolAfterPaste", True) Then toolbar_Toolbox.SelectNewTool NAV_MOVE
            
        End If
        
    Else
        PDMsgBox "The clipboard is empty or it does not contain a valid picture format.  Please copy a valid image onto the clipboard and try again.", vbExclamation Or vbOKOnly, "Error"
    End If
    
End Function

'If the clipboard contains internal PD-format data (most commonly a bare DIB, but also potentially a
' layer header and/or layer vector contents), you can call this function to initiate a "paste" command
' using the internal data as a source.  The parameter "srcIsMeantAsLayer" controls whether the clipboard
' data is loaded as a new image, or as a new layer in an existing image.
'
'RETURNS: TRUE if successful; FALSE otherwise.
Private Function ClipboardPaste_InternalData(ByVal srcIsMeantAsLayer As Boolean, Optional ByRef pasteToThisDIBInstead As pdDIB = Nothing) As Boolean
    
    'Unfortunately, a lot of things can go wrong when pasting custom image data, so we assume failure by default.
    ClipboardPaste_InternalData = False
    
    'See if a previous stash file exists
    If (LenB(m_StashFile) <> 0) Then
        If Files.FileExists(m_StashFile) Then
            
            PDDebug.LogAction "Loading internal PD clipboard data now..."
            
            'We've previously saved a valid stash file.  Attempt to load it as a new file.
            Dim sTitle As String
            sTitle = g_Language.TranslateMessage("Clipboard Image")
            sTitle = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
            
            'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
            If (Not pasteToThisDIBInstead Is Nothing) Then
                ClipboardPaste_InternalData = Loading.QuickLoadImageToDIB(m_StashFile, pasteToThisDIBInstead, False, False)
            Else
                If srcIsMeantAsLayer Then
                    
                    'If a layer header was saved out to file, try to use it as the paste source
                    ' (instead of creating a generic new layer object - this will result in a copy+paste operation
                    ' that preserves attributes like layer blend mode, position, rotation, text/vector contents, etc)
                    Dim useStashedLayerHeader As Boolean
                    useStashedLayerHeader = (LenB(m_StashFileLayer) <> 0)
                    If useStashedLayerHeader Then useStashedLayerHeader = Files.FileExists(m_StashFileLayer)
                    If useStashedLayerHeader Then useStashedLayerHeader = Files.FileLoadAsString(m_StashFileLayer, m_StashedLayer)
                    
                    'Only attempt to use a stashed layer header if all previous validation steps were successful.
                    If useStashedLayerHeader Then
                        
                        PDDebug.LogAction "Pasting internal layer source..."
                        
                        'Ask the parent pdImage to create a blank, new layer object
                        Dim newLayerID As Long
                        newLayerID = PDImages.GetActiveImage.CreateBlankLayer()
                        
                        'See if we stashed raster or vector data.  (This affects how we initialize the layer's contents -
                        ' from pixel buffer or manually, by reading text-based vector data and generating objects accordingly.)
                        Dim useVectorData As Boolean
                        useVectorData = False
                        
                        Dim cSerialize As pdSerialize
                        Set cSerialize = New pdSerialize
                        cSerialize.SetParamString m_StashedLayer
                        
                        Dim srcLayerType As PD_LayerType
                        srcLayerType = Layers.GetLayerTypeIDFromString(cSerialize.GetString("type", Layers.GetLayerTypeStringFromID(PDL_Image), True))
                        
                        If (srcLayerType <> PDL_Image) Then
                            useVectorData = Files.FileExists(m_StashFileVector)
                            If useVectorData Then useVectorData = Files.FileLoadAsString(m_StashFileVector, m_StashedVector)
                        End If
                        
                        'useVectorData will now be TRUE iff the copied layer was a text/vector layer,
                        ' *and* we stashed vector data successfully.
                        
                        'Create the new layer using the appropriate layer type
                        PDImages.GetActiveImage.GetLayerByID(newLayerID).InitializeNewLayer srcLayerType
                        
                        'Initialize the layer using the cached layer header.  (This preserves all editable layer attributes.)
                        PDImages.GetActiveImage.GetLayerByID(newLayerID).CreateNewLayerFromXML m_StashedLayer, newLayerID, True
                        
                        'If this is a vector layer, initialize it from text-based vector data
                        If useVectorData Then useVectorData = PDImages.GetActiveImage.GetLayerByID(newLayerID).SetVectorDataFromXML(m_StashedVector)
                        
                        'If the vector creation was successful, no further work is required.
                        ' (If vector creation was *not* successful - which is unexpected and shouldn't happen -
                        '  fall back to the cached raster data instead.)
                        If useVectorData Then
                            PDDebug.LogAction "Pasting used vector data."
                            ClipboardPaste_InternalData = True
                        Else
                            
                            PDDebug.LogAction "Paste will use raster data..."
                            
                            'Load the stashed pixel data into a standalone DIB
                            Dim tmpDIB As pdDIB
                            Set tmpDIB = New pdDIB
                            
                            If PDImages.GetActiveImage.GetLayerByID(newLayerID).AffineTransformsActive(True) Then
                                ClipboardPaste_InternalData = Loading.QuickLoadImageToDIB(m_StashFileLayerAffineOrig, tmpDIB, False, False)
                            Else
                                ClipboardPaste_InternalData = Loading.QuickLoadImageToDIB(m_StashFile, tmpDIB, False, False)
                            End If
                            
                            If ClipboardPaste_InternalData Then
                                
                                'Forcibly convert the new layer to 32bpp
                                ' (failsafe only; it should already be in 32-bpp mode from the loader)
                                If (tmpDIB.GetDIBColorDepth <> 32) Then tmpDIB.ConvertTo32bpp
                                
                                'Replace the new layer's backing surface with the temporary DIB
                                PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerDIB tmpDIB
                                
                            Else
                                PDDebug.LogAction "WARNING!  ClipboardPaste_InternalData failed to load the stash file.  Paste abandoned."
                            End If
                            
                        End If
                        
                        'Because this layer has been pasted with all custom parameters intact - including position! -
                        ' we now need to determine if the layer even appears within its new image container.
                        Layers.EnsureLayerInbounds newLayerID
                        
                        'Notify the parent image of these changes, as it needs to generate a new composite image
                        PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
                        
                    'If a layer header *wasn't* cached, we are probably copying from a non-layer source (like a selection).
                    ' Simply load the cached raster data as a new, standalone layer.
                    Else
                        m_LayerCounter = m_LayerCounter + 1
                        ClipboardPaste_InternalData = Layers.LoadImageAsNewLayer(False, m_StashFile, g_Language.TranslateMessage("Clipboard Image") & " " & CStr(m_LayerCounter), False, False)
                    End If
                    
                Else
                    ClipboardPaste_InternalData = Loading.LoadFileAsNewImage(m_StashFile, sTitle, False)
                End If
            End If
            
            'If we made it all the way here, the load was (probably?) successful
            If ClipboardPaste_InternalData Then Message "Clipboard data imported successfully "
            
        Else
            PDDebug.LogAction "WARNING!  ClipboardPaste_InternalData failed to find a stash file.  Paste abandoned."
        End If
    Else
        PDDebug.LogAction "WARNING!  ClipboardPaste_InternalData failed, because a stash file hasn't been created yet.  Paste abandoned."
    End If
    
End Function

'If the clipboard contains custom-format image data (most commonly PNG or JPEG), you can call this function
' to initiate a "paste" command using the custom image data as a source.  The parameter "srcIsMeantAsLayer"
' controls whether the clipboard data is loaded as a new image, or as a new layer in an existing image.
'
'RETURNS: TRUE if successful; FALSE otherwise.
Private Function ClipboardPaste_CustomImageFormat(ByVal clipboardFormatName As String, ByVal srcIsMeantAsLayer As Boolean, Optional ByVal tmpFileExtension As String = "tmp", Optional ByRef pasteToThisDIBInstead As pdDIB = Nothing) As Boolean
    
    'Unfortunately, a lot of things can go wrong when pasting custom image data, so we assume failure by default.
    ClipboardPaste_CustomImageFormat = False
    
    'All paste operations use a few consistent variables
    
    'Raw retrieval storage variables
    Dim clipFormatID As Long, rawClipboardData() As Byte
    
    'Temporary file for storing the clipboard data.  (This lets us use PD's central image load function.)
    Dim tmpClipboardFile As String
    
    'Verify that the requested data is actually available.  (Hopefully the caller already checked this, but you never know...)
    If m_Clipboard.DoesClipboardHaveFormatName(clipboardFormatName) Then
        
        PDDebug.LogAction "ClipboardPaste_CustomImageFormat() will now attempt to load " & clipboardFormatName & " from the clipboard..."
        
        'Because custom-format image data can be registered by many programs, retrieve this image format's unique ID now.
        clipFormatID = m_Clipboard.GetFormatIDFromName(clipboardFormatName)
        m_ClipboardInfo.pdci_CurrentFormat = clipFormatID
        m_ClipboardInfo.pdci_OriginalFormat = clipFormatID
        
        'Pull the data into a local array
        If m_Clipboard.GetClipboardBinaryData(clipFormatID, rawClipboardData) Then
            
            'Dump the data out to file
            tmpClipboardFile = UserPrefs.GetTempPath & "PDClipboard." & tmpFileExtension
            If Files.FileCreateFromByteArray(rawClipboardData, tmpClipboardFile) Then
                
                'We no longer need our local copy of the clipboard data
                Erase rawClipboardData
                
                'We can now use the standard image load routine to import the temporary file.  Because we don't want the
                ' load function to use the temporary file name as the image name, we manually supply a filename to suggest
                ' if the user eventually tries to save the file.
                Dim sTitle As String
                sTitle = g_Language.TranslateMessage("Clipboard Image")
                sTitle = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
                
                'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                If (Not pasteToThisDIBInstead Is Nothing) Then
                    ClipboardPaste_CustomImageFormat = Loading.QuickLoadImageToDIB(m_StashFile, pasteToThisDIBInstead, False, False)
                Else
                    If srcIsMeantAsLayer Then
                        m_LayerCounter = m_LayerCounter + 1
                        ClipboardPaste_CustomImageFormat = Layers.LoadImageAsNewLayer(False, tmpClipboardFile, g_Language.TranslateMessage("Clipboard Image") & " " & CStr(m_LayerCounter), False, False)
                    Else
                        ClipboardPaste_CustomImageFormat = LoadFileAsNewImage(tmpClipboardFile, sTitle, False)
                    End If
                End If
                
                'Be polite and remove the temporary file
                Files.FileDeleteIfExists tmpClipboardFile
                
                'If we made it all the way here, the load was (probably?) successful
                If ClipboardPaste_CustomImageFormat Then Message "Clipboard data imported successfully "
                
            Else
                PDDebug.LogAction "WARNING!  Clipboard image data (probably PNG) could not be written to a temp file."
            End If
            
        Else
            PDDebug.LogAction "WARNING!  Clipboard.GetBinaryData failed on custom image data (probably PNG).  Special paste action abandoned."
        End If
        
    Else
        PDDebug.LogAction "WARNING!  ClipboardPaste_CustomImageFormat was called, but the requested data doesn't exist on the clipboard."
    End If
    
End Function

'If the clipboard contains HTML text (presumably copied from a web browser), you can call this function to initiate a "paste" command
' using the HTML text as a source.  When copying an image from the web, most web browsers will include a link to the original image
' on the clipboard; we prefer to download this vs grabbing the actual image bits, as we provide much more comprehensive handling for
' things like metadata, special PNG chunks, ICC profiles, and more.
'Also, the parameter "srcIsMeantAsLayer" controls whether the clipboard data is loaded as a new image, or as a new layer in the
' active image.
'
'RETURNS: TRUE if successful; FALSE otherwise.
Private Function ClipboardPaste_HTML(ByVal srcIsMeantAsLayer As Boolean, Optional ByRef pasteToThisDIBInstead As pdDIB = Nothing) As Boolean
    
    'Unfortunately, a lot of things can go wrong when pasting custom image data, so we assume failure by default.
    ClipboardPaste_HTML = False
    
    'Verify that the requested data is actually available.  (Hopefully the caller already checked this, but you never know...)
    If m_Clipboard.DoesClipboardHaveHTML() Then
        
        'HTML handling requires no special behavior on the part of external load functions, so we mark the module-level tracker as blank
        m_ClipboardInfo.pdci_CurrentFormat = 0
        m_ClipboardInfo.pdci_OriginalFormat = 0
        
        PDDebug.LogAction "ClipboardPaste_HTML() will now attempt to find valid image HTML on the clipboard..."
        
        'Pull the HTML data into a local string
        Dim htmlString As String
        If m_Clipboard.GetClipboardHTML(htmlString) Then
            
            'Look for an image reference within the HTML snippet
            If (InStr(1, UCase$(htmlString), "<IMG ", vbBinaryCompare) > 0) Then
            
                'Retrieve the full image path, which will be between the first set of quotation marks following the
                ' "<img src=" statement in the HTML snippet.
                Dim vbQuoteMark As String
                vbQuoteMark = """"
                
                'Parse out the URL between the img src quotes
                Dim urlStart As Long, urlEnd As Long
                urlStart = InStr(1, UCase$(htmlString), "<IMG ", vbBinaryCompare)
                If (urlStart > 0) Then urlStart = InStr(urlStart, UCase$(htmlString), "SRC=", vbBinaryCompare)
                If (urlStart > 0) Then urlStart = InStr(urlStart, htmlString, vbQuoteMark, vbBinaryCompare) + 1
                
                'The magic number 6 below is calculated as the length of (src="), + 1 to advance to the
                ' character immediately following the quotation mark.
                If (urlStart > 0) Then urlEnd = InStr(urlStart + 6, htmlString, vbQuoteMark, vbBinaryCompare)
                
                'As a failsafe, make sure a valid URL was actually found
                If (urlStart > 0) And (urlEnd > 0) Then
                    
                    ClipboardPaste_HTML = ClipboardPaste_WellFormedURL(Mid$(htmlString, urlStart, urlEnd - urlStart), srcIsMeantAsLayer, True, pasteToThisDIBInstead)
                
                'An image tag was found, but a parsing error occurred when trying to strip out the source URL.  This is okay;
                ' exit immediately without raising any errors.
                Else
                    PDDebug.LogAction "Clipboard.GetClipboardHTML was successful and an image URL was located, but a parsing error occurred."
                End If
                
            'No image tag found, which is fine; exit immediately without raising any errors.
            Else
                PDDebug.LogAction "Clipboard.GetClipboardHTML was successful, but no image URL found."
            End If
            
        Else
            PDDebug.LogAction "WARNING!  Clipboard.GetClipboardHTML failed.  Special paste action abandoned."
        End If
        
    Else
        PDDebug.LogAction "WARNING!  ClipboardPaste_HTML was called, but HTML data doesn't exist on the clipboard."
    End If
    
End Function

'If one or more files exist on the clipboard, attempt to paste them all.
Private Function ClipboardPaste_ListOfFiles(ByVal srcIsMeantAsLayer As Boolean, Optional ByRef pasteToThisDIBInstead As pdDIB = Nothing) As Boolean
    
    'Make sure files actually exist on the clipboard
    If m_Clipboard.DoesClipboardHaveFiles() Then
        
        PDDebug.LogAction "ClipboardPaste_ListOfFiles() will now attempt to load one or more files from the clipboard..."
        
        'File lists require no special behavior on the part of external load functions, so we mark the module-level tracker as blank
        m_ClipboardInfo.pdci_CurrentFormat = CF_HDROP
        m_ClipboardInfo.pdci_OriginalFormat = CF_HDROP
        
        Dim listOfFiles As pdStringStack, numOfFiles As Long, tmpFilename As String
        If m_Clipboard.GetFileList(listOfFiles, numOfFiles) Then
            
            'We will report success if at least one file loads successfully
            ClipboardPaste_ListOfFiles = False
            
            'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
            If srcIsMeantAsLayer Then
                Do While listOfFiles.PopString(tmpFilename)
                    If (LenB(tmpFilename) <> 0) Then
                        ClipboardPaste_ListOfFiles = ClipboardPaste_ListOfFiles Or Layers.LoadImageAsNewLayer(False, tmpFilename, , False, False)
                    End If
                Loop
            Else
                ClipboardPaste_ListOfFiles = Loading.LoadMultipleImageFiles(listOfFiles)
            End If
            
        Else
            PDDebug.LogAction "WARNING!  ClipboardPaste_ListOfFiles couldn't retrieve a valid file list from pdClipboard."
        End If
        
    Else
        PDDebug.LogAction "WARNING!  ClipboardPaste_ListOfFiles was called, but no file paths exist on the clipboard."
    End If
        
End Function

'If the clipboard contains any valid metafile format, you can call this function to initiate a "paste" command.
' The parameter "srcIsMeantAsLayer" controls whether the clipboard data is loaded as a new image,
' or as a new layer in an existing image.
'
'RETURNS: TRUE if successful; FALSE otherwise.
Private Function ClipboardPaste_Metafile(ByVal srcIsMeantAsLayer As Boolean, Optional ByRef pasteToThisDIBInstead As pdDIB = Nothing) As Boolean
    
    'Unfortunately, a lot of things can go wrong when pasting custom image data, so we assume failure by default.
    ClipboardPaste_Metafile = False
    
    'Temporary file for storing the clipboard data.  (This lets us use PD's central image load function.)
    Dim tmpClipboardFile As String
    
    'Verify that the requested data is actually available.  (Hopefully the caller already checked this, but you never know...)
    'Metafiles show up in a few different formats.  We want to use EMFs, if possible.
    If m_Clipboard.DoesClipboardHaveFormatID(CF_ENHMETAFILE) Then
        
        PDDebug.LogAction "ClipboardPaste_Metafile() will now attempt to load an EMF from the clipboard..."
        
        'Clipboard metafiles only store a *handle* to the source metafile.  They do not store the actual metafile data,
        ' which means we have to manually grab the handle and write the data out to file.
        Dim emfHandle As Long
        emfHandle = m_Clipboard.GetClipboardMemoryHandle(CF_ENHMETAFILE)
        
        If (emfHandle <> 0) Then
        
            PDDebug.LogAction "Clipboard metafile handle retrieved successfully.  Writing to temp file..."
            tmpClipboardFile = UserPrefs.GetTempPath & "PDClipboard.emf"
            
            Dim tmpEmfCopy As Long
            tmpEmfCopy = CopyEnhMetaFileW(emfHandle, StrPtr(tmpClipboardFile))
            
            'When copying to file, Windows (for whatever reason) creates a dupliate EMF handle.  We are responsible
            ' for freeing this handle.
            If (tmpEmfCopy <> 0) Then
                
                DeleteEnhMetaFile tmpEmfCopy
                
                'We can now use the standard image load routine to import the temporary file.  Because we don't want the
                ' load function to use the temporary file name as the image name, we manually supply a filename to suggest
                ' if the user eventually tries to save the file.
                Dim sTitle As String
                sTitle = g_Language.TranslateMessage("Clipboard Image")
                sTitle = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
                
                'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                If (Not pasteToThisDIBInstead Is Nothing) Then
                    ClipboardPaste_Metafile = Loading.QuickLoadImageToDIB(tmpClipboardFile, pasteToThisDIBInstead, False, False)
                Else
                    If srcIsMeantAsLayer Then
                        m_LayerCounter = m_LayerCounter + 1
                        ClipboardPaste_Metafile = Layers.LoadImageAsNewLayer(False, tmpClipboardFile, g_Language.TranslateMessage("Clipboard Image") & " " & CStr(m_LayerCounter), False, False)
                    Else
                        ClipboardPaste_Metafile = LoadFileAsNewImage(tmpClipboardFile, sTitle, False)
                    End If
                End If
                
                'Be polite and remove the temporary file
                Files.FileDeleteIfExists tmpClipboardFile
                
                'If we made it all the way here, the load was (probably?) successful
                If ClipboardPaste_Metafile Then Message "Clipboard data imported successfully "
                
            Else
                PDDebug.LogAction "WARNING!  Clipboard metafile data could not be written to a temp file."
            End If
            
        Else
            PDDebug.LogAction "WARNING!  Clipboard.GetBinaryData failed on metafile data.  Special paste action abandoned."
        End If
        
    Else
        PDDebug.LogAction "WARNING!  ClipboardPaste_Metafile was called, but the requested data doesn't exist on the clipboard."
    End If
    
End Function

'If the clipboard contains text, try to find an image path or URL that we can use.
Private Function ClipboardPaste_TextSource(ByVal srcIsMeantAsLayer As Boolean, Optional ByRef pasteToThisDIBInstead As pdDIB = Nothing) As Boolean
    
    ClipboardPaste_TextSource = False
    
    'Make sure text actually exists on the clipboard
    If m_Clipboard.DoesClipboardHaveText() Then
        
        PDDebug.LogAction "ClipboardPaste_TextSource() will now parse clipboard text, looking for image sources..."
        
        'Text requires no special behavior on the part of external load functions, so we mark the module-level tracker as blank
        m_ClipboardInfo.pdci_CurrentFormat = 0
        m_ClipboardInfo.pdci_OriginalFormat = 0
        
        Dim clipText As String
        If m_Clipboard.GetClipboardText(clipText) Then
            
            'As reported by @manfromarce on GitHub (https://github.com/tannerhelland/PhotoDemon/discussions/421#discussioncomment-3476307)
            ' Windows Explorer provides a "Copy as path" context menu option that we should support - however, this option
            ' places double-quotes around the file path, which breaks our file path detector.
            
            'So look for and remove quotes around the target text, if any.
            clipText = Trim$(clipText)
            If (Left$(clipText, 1) = """") Then clipText = Right$(clipText, Len(clipText) - 1)
            If (Right$(clipText, 1) = """") Then clipText = Left$(clipText, Len(clipText) - 1)
            
            'Now we will test the text for various supportable types, starting with URLs
            If (Strings.StringsEqualLeft(clipText, "http", True) Or Strings.StringsEqualLeft(clipText, "ftp", True)) Then
                Message "Image URL found.  Attempting to download..."
                ClipboardPaste_TextSource = ClipboardPaste_WellFormedURL(clipText, srcIsMeantAsLayer)
                
            'If this doesn't look like a URL, see if it's a file path instead
            Else
                
                'If the text matches a local file path, try to load it.
                If Files.FileExists(clipText) Then
                    
                    If (Not pasteToThisDIBInstead Is Nothing) Then
                        ClipboardPaste_TextSource = Loading.QuickLoadImageToDIB(clipText, pasteToThisDIBInstead, False, False)
                    Else
                        If srcIsMeantAsLayer Then
                            ClipboardPaste_TextSource = Layers.LoadImageAsNewLayer(False, clipText, vbNullString, False, False)
                        Else
                            ClipboardPaste_TextSource = LoadFileAsNewImage(clipText)
                        End If
                    End If
                    
                End If
                
            End If
            
        Else
            PDDebug.LogAction "WARNING!  ClipboardPaste_TextSource couldn't retrieve actual text from pdClipboard."
        End If
        
    Else
        PDDebug.LogAction "WARNING!  ClipboardPaste_TextSource was called, but no text exists on the clipboard."
    End If
        
End Function

'Helper function that SHOULD NOT BE CALLED DIRECTLY.  Only other ClipboardPaste_* variants are able to safely use this function.
' Returns: TRUE if an image was successfully downloaded and loaded to the pdImages collection.  FALSE if failure occurred.
Private Function ClipboardPaste_WellFormedURL(ByVal srcURL As String, ByVal srcIsMeantAsLayer As Boolean, Optional ByVal suppressErrorMsgs As Boolean = False, Optional ByRef pasteToThisDIBInstead As pdDIB = Nothing) As Boolean
    
    'This function assumes the source URL is both absolute and well-formed
    Message "Image URL found.  Attempting to download..."
                    
    Dim tmpDownloadFile As String
    tmpDownloadFile = Web.DownloadURLToTempFile(srcURL, suppressErrorMsgs)
    
    'If the download was successful, we can now use the standard image load routine to import the temporary file
    If (Files.FileLenW(tmpDownloadFile) <> 0) Then
        
        'Additional file information variables, which we pass to the central load function to let it know that this is only a temp file,
        ' and it should use these hint values instead of assuming normal image load behavior.
        Dim tmpFilename As String
        tmpFilename = Files.FileGetName(tmpDownloadFile, True)
        
        'Depending on the request, load the clipboard data as a new image or as a new layer in the current image.
        ' (Note that we also suspend load errors, because if the URL was bad, the generic "clipboard doesn't have
        ' useable data" message covers all our bases here.)
        If (Not pasteToThisDIBInstead Is Nothing) Then
            ClipboardPaste_WellFormedURL = Loading.QuickLoadImageToDIB(tmpDownloadFile, pasteToThisDIBInstead, False, False)
        Else
            If (PDImages.IsImageActive() And srcIsMeantAsLayer) Then
                ClipboardPaste_WellFormedURL = Layers.LoadImageAsNewLayer(False, tmpDownloadFile, tmpFilename, False, False)
            Else
                ClipboardPaste_WellFormedURL = LoadFileAsNewImage(tmpDownloadFile, tmpFilename, False, vbYes)
            End If
        End If
        
        'Delete the temporary file
        Files.FileDeleteIfExists tmpDownloadFile
        
        If (Not suppressErrorMsgs) Then
            If ClipboardPaste_WellFormedURL Then
                Message "Image imported successfully "
            Else
                Message "Image download failed.  Please copy a valid image URL and try again."
            End If
        End If
        
        'Check for load failure.  If the most recent pdImages object is inactive, it's a safe assumption that
        ' the load operation failed.  (This isn't foolproof, especially if the user loads a ton of images,
        ' and subsequently unloads images in an arbitrary order - but given the rarity of that situation,
        ' I'm content to use this simplified technique for predicting failure.)
        If ((pasteToThisDIBInstead Is Nothing) And PDImages.IsImageActive(PDImages.GetActiveImageID())) Then
            ClipboardPaste_WellFormedURL = PDImages.GetActiveImage.IsActive()
        End If
    
    'If the download failed, let the user know that hey, at least we tried.
    Else
        Message "Image download failed.  Please copy a valid image URL and try again."
    End If
    
End Function

'If the clipboard contains bitmap-format image data (or by extension, DIB or DIBv5), you can call this function to initiate a "paste" command.
' The function will automatically determine the best format for pasting.  The parameter "srcIsMeantAsLayer" controls whether the clipboard data
' is loaded as a new image, or as a new layer in an existing image.
'
'RETURNS: TRUE if successful; FALSE otherwise.
Private Function ClipboardPaste_BitmapImage(ByVal srcIsMeantAsLayer As Boolean, Optional ByRef pasteToThisDIBInstead As pdDIB = Nothing) As Boolean
        
    'Unfortunately, a lot of things can go wrong when pasting bitmaps, so we assume failure by default.
    ClipboardPaste_BitmapImage = False
    
    'Verify that the requested data is actually available.  (Hopefully the caller already checked this, but you never know...)
    If m_Clipboard.DoesClipboardHaveBitmapImage() Then
        
        'Next, we want to sort handling by the "priority" bitmap format.  This is the format the caller actually placed on the clipboard
        ' (vs a variant that Windows auto-created to simplify handling).
        Dim priorityFormat As PredefinedClipboardFormatConstants
        priorityFormat = m_Clipboard.GetPriorityBitmapFormat()
        
        'Bitmap formats may require special behavior on the part of external load functions, so it's important that we accurately
        ' mark the module-level tracker with both the current format (what we retrieved from the clipboard; this may have been
        ' auto-generated by Windows), and the original format the caller placed on the clipboard.
        m_ClipboardInfo.pdci_OriginalFormat = priorityFormat
        
        'If DIBv5 is the format the caller actually placed on the clipboard, retrieve it first.  Otherwise, use the CF_DIB data.
        ' (Ignore CF_BITMAP for now, as it would require specialized handling.)
        Dim rawClipboardData() As Byte, successfulExtraction As Boolean
        If (priorityFormat = CF_DIBV5) Then
            PDDebug.LogAction "DIBv5 selected as priority retrieval format."
            successfulExtraction = m_Clipboard.GetClipboardBinaryData(CF_DIBV5, rawClipboardData)
            m_ClipboardInfo.pdci_CurrentFormat = CF_DIBV5
        Else
            PDDebug.LogAction "Generic DIB selected as priority retrieval format."
            successfulExtraction = m_Clipboard.GetClipboardBinaryData(CF_DIB, rawClipboardData)
            m_ClipboardInfo.pdci_CurrentFormat = CF_DIB
        End If
        
        'If the extraction was successful, we can use similar handling for both cases
        If successfulExtraction Then
        
            'Perform some failsafe validation on the DIB header
            Dim dibHeaderOkay As Boolean
            
            'First, make sure we have at least 40 bytes of data to work with.  (Anything smaller than this
            ' and we can't even retrieve a header!)
            dibHeaderOkay = (UBound(rawClipboardData) >= 40)
            
            'If we have at least 40 bytes of data, copy them into a default BITMAPINFOHEADER.
            ' This struct is shared between regular DIBs and v5 DIBs.
            Dim bmpHeader As BITMAPINFOHEADER, bmpV5Header As BITMAPV5HEADER
            If dibHeaderOkay Then
                
                'Retrieve a copy of the bitmap's header in standard, 40-byte format.
                ' This reports useful values like width, height, and color depth.
                CopyMemoryStrict VarPtr(bmpHeader), VarPtr(rawClipboardData(0)), LenB(bmpHeader)
                
                'Validate the header size; it must match a default DIB header, or a v5 DIB header
                dibHeaderOkay = (bmpHeader.biSize = LenB(bmpHeader)) Or (bmpHeader.biSize = LenB(bmpV5Header))
                
                'If a v5 header is present, retrieve it as well
                If (priorityFormat = CF_DIBV5) And (bmpHeader.biSize = LenB(bmpV5Header)) And (UBound(rawClipboardData) > LenB(bmpV5Header)) Then
                    
                    CopyMemoryStrict VarPtr(bmpV5Header), VarPtr(rawClipboardData(0)), LenB(bmpV5Header)
                    
                    'Track some v5 header data at module-level; external functions may request copies of this,
                    ' to know how to handle alpha.
                    m_ClipboardInfo.pdci_DIBv5AlphaMask = bmpV5Header.biAlphaMask
                    
                End If
                
                'If the header size checks out, validate width/height next
                If dibHeaderOkay Then
                
                    With bmpHeader
                    
                        'Width must be positive
                        If (.biWidth < 0) Then dibHeaderOkay = False
                        
                        'For performance reasons, restrict sizes to 2 ^ 16 in either dimension.
                        ' (This metric is also used by Chrome and Firefox.)
                        If (.biWidth > (2 ^ 16)) Or (Abs(.biHeight) > (2 ^ 16)) Then dibHeaderOkay = False
                        
                        'Check for invalid bit-depths.
                        If (.biBitCount <> 1) And (.biBitCount <> 4) And (.biBitCount <> 8) And (.biBitCount <> 16) And (.biBitCount <> 24) And (.biBitCount <> 32) Then dibHeaderOkay = False
                        
                        'Check for invalid compression sub-types
                        If (.biCompression > BC_BITFIELDS) Then dibHeaderOkay = False
                    
                    End With
                    
                End If
                
                'We've now performed pretty reasonable header validation.  If the header passed, proceed with parsing.
                If dibHeaderOkay Then
                    
                    'Prepare a temporary DIB to receive a 24 or 32-bit copy of the clipboard data.
                    Dim tmpDIB As pdDIB
                    Set tmpDIB = New pdDIB
                    
                    'See if a 24 or 32-bit destination image is required
                    If (bmpHeader.biBitCount = 32) Or ((bmpHeader.biSize = LenB(bmpV5Header)) And (bmpV5Header.biAlphaMask <> 0)) Then
                        tmpDIB.CreateBlank bmpHeader.biWidth, Abs(bmpHeader.biHeight), 32, 0, 0
                    Else
                        tmpDIB.CreateBlank bmpHeader.biWidth, Abs(bmpHeader.biHeight), 24, 0, 0
                    End If
                    
                    'Calculate the offset required to access the pixel data.  (This value is required by the BMP file format, which PD
                    ' uses as a quick intermediary format.)  Note that some offset calculations only apply to the v5 version of the header.
                    Dim pixelOffset As Long
                    
                    With bmpHeader
                        
                        'Always count the header size in the offset
                        pixelOffset = .biSize
                        
                        'If a color table is included, add it to the offset
                        If (.biClrUsed > 0) Then
                            pixelOffset = pixelOffset + .biClrUsed * 4
                        Else
                            If (.biBitCount <= 8) Then pixelOffset = pixelOffset + 4 * (2 ^ .biBitCount)
                        End If
                        
                        'Bitfields are optional with certain bit-depths; if bitfields are specified, add them too.
                        If (.biCompression = 3) Then
                            If (.biBitCount = 16) Or (.biBitCount = 32) Then pixelOffset = pixelOffset + 12
                        End If
                        
                    End With
                    
                    'v5 of the BMP header allows for ICC profiles.  These are supposed to be stored AFTER pixel data,
                    ' but some software is written by idiots (hi!), so perform a failsafe check for out-of-place profiles.
                    If (priorityFormat = CF_DIBV5) And (bmpV5Header.biProfileData <= pixelOffset) Then pixelOffset = pixelOffset + bmpV5Header.biProfileSize
                                        
                    'We now know enough to create a temporary BMP file as a placeholder for the clipboard data.
                    
                    'Place the temporary file in inside the program-specified temp path
                    Dim tmpClipboardFile As String
                    tmpClipboardFile = UserPrefs.GetTempPath & "PDClipboard.tmp"
                    
                    'pdFSO is used to ensure Unicode subfolder compatibility
                    Files.FileDeleteIfExists tmpClipboardFile
                    
                    'Populate the BMP file header; it's a simple 14-byte, unchanging struct that requires only a magic number,
                    ' a total filesize, and an offset that points at the pixel bits (NOT the BMP file header, or the embedded
                    ' DIB header - the actual pixel bits).
                    Dim bmpFileHeader As BITMAPFILEHEADER
                    With bmpFileHeader
                        .Type = &H4D42
                        .Size = (UBound(rawClipboardData) + 1) + 14
                        .OffBits = pixelOffset + 14
                    End With
                    
                    Dim cFile As pdFSO
                    Set cFile = New pdFSO
                    
                    Dim hFile As Long
                    If cFile.FileCreateAppendHandle(tmpClipboardFile, hFile) Then
                        
                        'To avoid automatic 4-byte struct alignment, we must write out the header manually.
                        cFile.FileWriteData hFile, VarPtr(bmpFileHeader.Type), 2&
                        cFile.FileWriteData hFile, VarPtr(bmpFileHeader.Size), 4&
            
                        Dim reservedBytes As Long
                        cFile.FileWriteData hFile, VarPtr(reservedBytes), 4&
                        cFile.FileWriteData hFile, VarPtr(bmpFileHeader.OffBits), 4&
                        
                        'Simply plop the clipboard data into place last, no changes required
                        cFile.FileWriteData hFile, VarPtr(rawClipboardData(0)), UBound(rawClipboardData) + 1
                        cFile.FileCloseHandle hFile
                        
                    End If
                        
                    'We can now use PD's standard image load routine to import the temporary file.  Because we don't want the
                    ' load function to use the temporary file name as the image name, we manually supply a filename to suggest
                    ' if the user eventually tries to save the file.
                    Dim sTitle As String
                    sTitle = g_Language.TranslateMessage("Clipboard Image")
                    sTitle = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
                    
                    'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                    If (Not pasteToThisDIBInstead Is Nothing) Then
                        ClipboardPaste_BitmapImage = Loading.QuickLoadImageToDIB(tmpClipboardFile, pasteToThisDIBInstead, False, False)
                    Else
                        If PDImages.IsImageActive() And srcIsMeantAsLayer Then
                            m_LayerCounter = m_LayerCounter + 1
                            ClipboardPaste_BitmapImage = Layers.LoadImageAsNewLayer(False, tmpClipboardFile, g_Language.TranslateMessage("Clipboard Image") & " " & CStr(m_LayerCounter), False, False)
                        Else
                            ClipboardPaste_BitmapImage = Loading.LoadFileAsNewImage(tmpClipboardFile, sTitle, False)
                        End If
                    End If
                    
                    'Once the load is complete, be polite and remove the temporary file
                    cFile.FileDelete tmpClipboardFile
                            
                    If ClipboardPaste_BitmapImage Then Message "Clipboard data imported successfully "
                    
                Else
                    PDDebug.LogAction "WARNING!  ClipboardPaste_BitmapImage failed because the DIB header failed validation.  Paste abandoned."
                End If
                
            Else
                PDDebug.LogAction "WARNING!  ClipboardPaste_BitmapImage failed because the DIB header is an invalid size.  Paste abandoned."
            End If
            
        Else
            PDDebug.LogAction "WARNING!  ClipboardPaste_BitmapImage failed to retrieve raw DIB data.  Paste abandoned."
        End If
        
    Else
        PDDebug.LogAction "WARNING!  ClipboardPaste_BitmapImage was called, but the requested data doesn't exist on the clipboard."
    End If
    
End Function

'Any droppable surface in the program can pass this function its DataObject to see if the dragged object can be successfully handled
' by PD.  Note that this function only returns whether the format is parseable by PD - an actual drop will need to occur before
' full validation is handled.  (This is necessary for things like URLs, which may take some time to validate.)
Friend Function IsObjectDragDroppable(ByRef Data As DataObject) As Boolean
    
    IsObjectDragDroppable = False
    
    'Make sure PD hasn't disabled program-wide drag/drop behavior (e.g. a modal form hasn't stolen focus, processing isn't occurring, etc)
    If g_AllowDragAndDrop Then
        If Data.GetFormat(vbCFFiles) Then IsObjectDragDroppable = True
        If Data.GetFormat(vbCFText) Then IsObjectDragDroppable = True
        If Data.GetFormat(vbCFBitmap) Then IsObjectDragDroppable = True
        If Data.GetFormat(CF_FileContents) Then IsObjectDragDroppable = True
        If Data.GetFormat(CF_FileGroupDescriptorW) Then IsObjectDragDroppable = True
    End If
    
End Function

'Because OLE drag/drop is so similar to clipboard functionality, I have included that functionality here.
' Data and Effect are passed as-is from the calling function, while intendedTargetIsLayer controls whether
' the image file(s) (if valid) should be loaded as new layers or new images.
Friend Function LoadImageFromDragDrop(ByRef Data As DataObject, ByRef Effect As Long, ByVal intendedTargetIsLayer As Boolean, Optional ByVal srcX As Long = LONG_MAX, Optional ByVal srcY As Long = LONG_MAX) As Boolean
    
    If (Data Is Nothing) Then Exit Function
    
    Dim dragWasSuccessful As Boolean
    dragWasSuccessful = False
    
    'When dragging from things like web browsers, there are a *ton* of potential formats we can use for retrieval:
    ' URL snippets, temporary cached copies, actual in-memory bitmaps.  If multiple formats are available, we'll
    ' suppress any intermediary error messages until the last possible format fails.  (Otherwise, if we raise a
    ' warning on the frist format, it's probably premature as we'll just grab a subsequent format.)
    Dim numSourcesAvailable As Long
    If Data.GetFormat(vbCFText) Then numSourcesAvailable = numSourcesAvailable + 1
    If Data.GetFormat(CF_FileContents) Then numSourcesAvailable = numSourcesAvailable + 1
    If Data.GetFormat(CF_FileGroupDescriptorW) Then numSourcesAvailable = numSourcesAvailable + 1
    If Data.GetFormat(vbCFFiles) Then numSourcesAvailable = numSourcesAvailable + 1
    If Data.GetFormat(vbCFBitmap) Then numSourcesAvailable = numSourcesAvailable + 1
    
    'Before calling individual load functions, convert the incoming x/y parameters to image coordinates.
    ' (Note that this is only useful if the target is a layer - when loading dropped files as new images,
    ' coordinates are ignored.)
    Dim newX As Long, newY As Long
    If intendedTargetIsLayer And (PDImages.GetNumOpenImages > 0) And (srcX <> LONG_MAX) And (srcY <> LONG_MAX) Then
        Dim tmpX As Double, tmpY As Double
        Drawing.ConvertCanvasCoordsToImageCoords FormMain.MainCanvas(0), PDImages.GetActiveImage(), srcX, srcY, tmpX, tmpY, True
        newX = Int(tmpX + 0.5)
        newY = Int(tmpY + 0.5)
    Else
        newX = LONG_MAX
        newY = LONG_MAX
    End If
    
    'First, look for text fragments that contain valid file or URL paths.  (These are extremely fast to query,
    ' and they're preferred with web browsers, as they generally contain a URL to the image in question,
    ' which we prefer to download manually so we can keep things like metadata intact.)
    If Data.GetFormat(vbCFText) And (Not dragWasSuccessful) Then
        dragWasSuccessful = DragDrop_TextData(Data, intendedTargetIsLayer, numSourcesAvailable > 1, newX, newY)
    End If
    
    'If we can't find text, look for file content or descriptor flags.  These are received when dragging from
    ' .zip files, and we can use them to manually reconstruct image files.
    If (Data.GetFormat(CF_FileContents) Or Data.GetFormat(CF_FileGroupDescriptorW)) And (Not dragWasSuccessful) Then
        dragWasSuccessful = DragDrop_ZippedFileList(Data, intendedTargetIsLayer, newX, newY)
    End If
    
    'If we're still here, look for formal file or file list objects.  These may be supplied by 3rd-party
    ' file explorer programs.
    If Data.GetFormat(vbCFFiles) And (Not dragWasSuccessful) Then
        dragWasSuccessful = DragDrop_FileList(Data, intendedTargetIsLayer, newX, newY)
    End If
    
    'If the data is not a file list, see if it's actual bitmap data.
    If Data.GetFormat(vbCFBitmap) And (Not dragWasSuccessful) Then
        dragWasSuccessful = DragDrop_BitmapData(Data, intendedTargetIsLayer, newX, newY)
    End If
    
    'If the drag was successful, and the result is a new layer, switch to the move/size tool
    ' (as the user probably needs to position the new layer)
    If (intendedTargetIsLayer And dragWasSuccessful) Then toolbar_Toolbox.SelectNewTool NAV_MOVE
    
    'If any of the supported formats worked, treat the drag/drop operation as successful
    LoadImageFromDragDrop = dragWasSuccessful
    
End Function

'This Unicode-friendly drag/drop interpreter comes courtesy LaVolpe, c/o
' http://cyberactivex.com/UnicodeTutorialVb.htm#Filenames_via_DragDrop_or_Paste (retrieved on 27/Feb/2016).
' IMPORTANT NOTE: this function has been modified for use inside PD.  If using it in your own project, I strongly recommend
' downloading the original version, as it includes additional helper code and explanations.
'
'Returns: TRUE if at least one file is contained in the passed OLEDragDrop_DataObject
Private Function DragDrop_FileList(ByRef OLEDragDrop_DataObject As DataObject, Optional ByVal loadAsLayers As Boolean = False, Optional ByVal srcX As Long = LONG_MAX, Optional ByVal srcY As Long = LONG_MAX) As Boolean
    
    'Unicode-enabled drag/drop code is used elsewhere in PD, so we use a central function for it
    Dim cFiles As pdStringStack
    If VBHacks.GetDragDropFileListW(OLEDragDrop_DataObject, cFiles) Then

        'The cFiles string stack now contains all the valid filenames from the dropped list.
        If (cFiles.GetNumOfStrings > 0) Then
            
            'Report success if at least one file loads successfully
            DragDrop_FileList = False
            
            If PDImages.IsImageActive() And loadAsLayers Then
                Dim i As Long
                For i = 0 To cFiles.GetNumOfStrings - 1
                    DragDrop_FileList = DragDrop_FileList Or Layers.LoadImageAsNewLayer(False, cFiles.GetString(i), , True, , srcX, srcY)
                Next i
            Else
                DragDrop_FileList = Loading.LoadMultipleImageFiles(cFiles)
            End If
            
        End If
        
    End If
    
End Function

'This Unicode-friendly drag/drop interpreter comes courtesy LaVolpe, c/o
' http://www.vbforums.com/showthread.php?637335-RESOLVED-Drag-amp-Drop-file-from-compressed-folder&p=3943417&viewfull=1#post3943417
' IMPORTANT NOTE: this function has been modified for use inside PD.  If using it in your own project, I strongly recommend
' downloading the original version, as it includes additional helper code and explanations.
'
'Returns: TRUE if at least one file is contained in the passed OLEDragDrop_DataObject
Private Function DragDrop_ZippedFileList(ByRef OLEDragDrop_DataObject As DataObject, Optional ByVal loadAsLayers As Boolean = False, Optional ByVal srcX As Long = LONG_MAX, Optional ByVal srcY As Long = LONG_MAX) As Boolean
    
    On Error GoTo DragDropZippedFilesFailed
    
    DragDrop_ZippedFileList = False
    
    'Note that PD will validate the data object before passing it, but better safe than sorry
    If (OLEDragDrop_DataObject Is Nothing) Then Exit Function
    If (Not OLEDragDrop_DataObject.GetFormat(CF_FileContents)) And (Not OLEDragDrop_DataObject.GetFormat(CF_FileGroupDescriptorW)) Then Exit Function
    
    'This function basically handles the task of hacking around the DataObject's VTable, and manually retrieving pointers
    ' to the original, unmodified Unicode filenames in the Data object.  We then overwrite the OUT filename strings with
    ' new Unicode-enabled pointers, so our receiving function can proceed identically regardless of the presence of
    ' non-ANSI chars.
    Dim fmtEtc As FORMATETC
    With fmtEtc
        .cfFormat = CF_FileGroupDescriptorW
        .lIndex = -1                  ' -1 means "we want everything"
        .TYMED = TYMED_HGLOBAL        ' TYMED_HGLOBAL means we want to use "hGlobal" as the transfer medium
        .dwAspect = DVASPECT_CONTENT  ' dwAspect is used to request extra metadata (like an icon representation) - we want the actual data
    End With

    'The IDataObject pointer appears 16 bytes past VB's DataObject
    Dim IID_IDataObject As Long
    CopyMemoryStrict VarPtr(IID_IDataObject), ObjPtr(OLEDragDrop_DataObject) + 16&, 4&
    
    'The objPtr of the IDataObject interface also tells us where the interface's VTable begins.  Since we know the
    ' VTable address and we know which function index we want, we can call it directly using the DispCallFunc.
    ' (You could also do this using a TLB, obviously.)  In particular, we want the GetData function which is #4 in
    ' the VTable, per http://msdn2.microsoft.com/en-us/library/ms688421.aspx
    
    'Next, we need to populate the input values required by the OLE API (http://msdn2.microsoft.com/en-us/library/ms221473.aspx)
    Dim pVartypes(0 To 3) As Integer, Vars(0 To 3) As Variant, pVars(0 To 3) As Long
    pVartypes(0) = vbLong: Vars(0) = VarPtr(fmtEtc): pVars(0) = VarPtr(Vars(0))
    
    Dim pMedium As STGMEDIUM, fMedium As STGMEDIUM
    pVartypes(1) = vbLong: Vars(1) = VarPtr(pMedium): pVars(1) = VarPtr(Vars(1))
    
    'Manually invoke the desired interface
    Dim varRtn As Variant
    If (DispCallFunc(IID_IDataObject, IDataObjVTable_GetData, CC_STDCALL, vbLong, 2, pVartypes(0), pVars(0), varRtn) = 0) Then
        
        'Make sure we received a non-null hGlobal pointer (nothing needs to be freed at this point, FYI)
        If (pMedium.Data = 0) Then Exit Function
        
        'pMedium.Data is an hGlobal handle that we can immediately lock and use
        Dim hGlobalPtr As Long
        hGlobalPtr = GlobalLock(pMedium.Data)
        
        'Technically we should never retrieve have received a null-pointer, but again, better safe than sorry
        If (hGlobalPtr <> 0) Then
            
            'Retrieve the number of entries (files and/or folders) in the drag
            Dim numOfFiles As Long
            CopyMemoryStrict VarPtr(numOfFiles), hGlobalPtr, 4&
            
            'We've been handed a list of file descriptors.  For each valid descriptor, we want to retrieve a matching
            ' file content object.  We'l request these as IStreams to simplify the creation of temporary files.
            fmtEtc.cfFormat = CF_FileContents
            fmtEtc.TYMED = TYMED_HGLOBAL
            Vars(1) = VarPtr(fMedium)
            
            'After the file+folder count comes the first file descriptor.
            hGlobalPtr = hGlobalPtr + 4&
            
            Dim fileDesc As FILEDESCRIPTOR
            Dim lLen As Long, fileHGlobalPtr As Long, tmpFilename As String, tmpBytes() As Byte
            Dim listOfFiles As pdStringStack: Set listOfFiles = New pdStringStack
            
            'We're now going to loop through each descriptor in turn, copying them out to temp files as we go
            Dim i As Long
            For i = 0 To numOfFiles - 1
            
                'Start by grabbing the file descriptor
                CopyMemoryStrict VarPtr(fileDesc), hGlobalPtr, LenB(fileDesc)
                
                'Pull out the filename
                lLen = lstrlenW(VarPtr(fileDesc.cFileName(0))) ' get length of filename
                If (lLen > 0) Then
                    tmpFilename = String$(lLen, 0)
                    CopyMemoryStrict StrPtr(tmpFilename), VarPtr(fileDesc.cFileName(0)), lLen * 2
                End If
                
                'In the future, it might be nice/clever to handle folders, so I've left this here, but for now, we ignore them
                If (fileDesc.dwFileAttributes And vbDirectory) = vbDirectory Then
                    'Create the folder here, potentially?
                Else
                    
                    'We're now going to raise a new GetData request, this time for the file's actual contents.
                    fmtEtc.lIndex = i
                    If (DispCallFunc(IID_IDataObject, IDataObjVTable_GetData, CC_STDCALL, vbLong, 2&, pVartypes(0), pVars(0), varRtn) = 0) Then
                        
                        'See if our request for an hGlobal pointer was honored
                        If (fMedium.TYMED = TYMED_HGLOBAL) Then
                            
                            'Save the file's contents to a temporary file, and note the filename (if successful)
                            fileHGlobalPtr = GlobalLock(fMedium.Data)
                            If Files.FileCreateFromPtr(fileHGlobalPtr, GlobalSize(fMedium.Data), UserPrefs.GetTempPath & tmpFilename, True, True) Then
                                listOfFiles.AddString UserPrefs.GetTempPath & tmpFilename
                            End If
                            
                            GlobalUnlock fMedium.Data
                            
                        'Depending on the size of the file, our request may get ignored and we'll be handed an IStream instead.
                        ' We can work with this if we have to.
                        ElseIf (fMedium.TYMED = TYMED_ISTREAM) Then
                            
                            'Make sure the stream has non-zero content length
                            If (fileDesc.nFileSizeLow > 0) Then
                            
                                'Save the file's contents to a temporary file, and note the filename (if successful)
                                VBHacks.ReadIStreamIntoVBArray fMedium.Data, tmpBytes, fileDesc.nFileSizeLow
                                If Files.FileCreateFromByteArray(tmpBytes, UserPrefs.GetTempPath & tmpFilename, True, True) Then
                                    listOfFiles.AddString UserPrefs.GetTempPath & tmpFilename
                                End If
                                
                            End If
                        
                        'Other returns are okay; we have fallback methods that deal with them successfully.
                        Else
                            'Debug.Print "WARNING!  DragDrop_ZippedFileList received " & fMedium.TYMED & " instead of an hGlobal or IStream; I can't handle that format."
                        End If
                        
                        'Conditionally release the stream (whether we do this or the caller does this is decided by
                        ' the caller, not us).
                        If (fMedium.pUnkForRelease = 0) Then ReleaseStgMedium VarPtr(fMedium)
                        
                    End If
                    
                End If
                
                'Advance to the next file descriptor
                hGlobalPtr = hGlobalPtr + LenB(fileDesc)
            
            Next i
            
            'We've got what we need from the hGlobal pointer, so go ahead and free it (if we're responsible for the drop -
            ' note that a NULL value for pUnkForRelease means that we must free the data; non-NULL means the caller will do it.)
            GlobalUnlock pMedium.Data
            If (pMedium.pUnkForRelease = 0) Then ReleaseStgMedium VarPtr(pMedium)
                
            'The listOfFiles string stack now contains all the valid filenames from the dropped data object.
            If (listOfFiles.GetNumOfStrings > 0) Then
                
                'Report success if at least one file loads correctly
                DragDrop_ZippedFileList = False
                
                'Load as layers...
                If PDImages.IsImageActive() And loadAsLayers Then
                
                    For i = 0 To listOfFiles.GetNumOfStrings - 1
                        If Files.FileExists(listOfFiles.GetString(i)) Then
                            DragDrop_ZippedFileList = DragDrop_ZippedFileList Or Layers.LoadImageAsNewLayer(False, listOfFiles.GetString(i), Files.FileGetName(listOfFiles.GetString(i), True), True, , srcX, srcY)
                            Files.FileDelete listOfFiles.GetString(i)
                        End If
                    Next i
                
                'Load as unique images...
                Else
                    
                    'Now we have to do something kinda ugly.  We don't want to load the bare filenames, because they are
                    ' just temp files - and we don't want to prompt the user to save stuff in the temp folder!
                    ' So instead, we'll pass each file individually to the single-load image function, and we'll manually
                    ' supply a custom title for each file.
                    Dim sTitle As String
                    Do While listOfFiles.PopString(tmpFilename)
                        If Files.FileExists(tmpFilename) Then
                            sTitle = Files.FileGetName(tmpFilename, True)
                            DragDrop_ZippedFileList = DragDrop_ZippedFileList Or LoadFileAsNewImage(tmpFilename, sTitle, False)
                            Files.FileDelete tmpFilename
                        End If
                    Loop
                    
                End If
                
            End If
            
        '/End non-zero hGlobal
        End If

    '/End DispCallFunc success
    End If
    
    Exit Function
    
DragDropZippedFilesFailed:
    PDDebug.LogAction "WARNING!  pdClipboardMain.DragDrop_ZippedFileList() experienced error #" & Err.Number & ": " & Err.Description
    
End Function

'Load dragged BMP data as a new image and/or layer.
Private Function DragDrop_BitmapData(ByRef OLEDragDrop_DataObject As DataObject, Optional ByVal loadAsLayers As Boolean = False, Optional ByVal srcX As Long = LONG_MAX, Optional ByVal srcY As Long = LONG_MAX) As Boolean
    
    On Error GoTo DragDropBMPFailed
    
    DragDrop_BitmapData = False
    
    'Note that PD will validate the data object before passing it, but better safe than sorry
    If (OLEDragDrop_DataObject Is Nothing) Then Exit Function
    If Not OLEDragDrop_DataObject.GetFormat(vbCFBitmap) Then Exit Function
        
    'It's hackish, but frankly the easiest way to grab the data is to simply copy it into a VB StdPicture container.
    Dim tmpPicture As StdPicture: Set tmpPicture = OLEDragDrop_DataObject.GetData(vbCFBitmap)
    
    'In the event that the bitmap contains alpha data, let's migrate it from that picture container to a pdDIB object
    Dim tmpDIB As pdDIB: Set tmpDIB = New pdDIB
    DIBs.CreateDIBFromStdPicture tmpDIB, tmpPicture
    Set tmpPicture = Nothing
    
    'Ask the DIB to write its contents to file in BMP format
    Dim tmpDragDropFile As String
    tmpDragDropFile = UserPrefs.GetTempPath & "PDDragDrop.tmp"
    tmpDIB.WriteToBitmapFile tmpDragDropFile
    Set tmpDIB = Nothing
    
    'We can now use PD's standard image load routine to import the temporary file.  Because we don't want the
    ' load function to use the temporary file name as the image name, we manually supply a filename to suggest
    ' if the user eventually tries to save the file.
    Dim sTitle As String
    sTitle = g_Language.TranslateMessage("Imported Image")
    sTitle = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
    
    'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
    If PDImages.IsImageActive() And loadAsLayers Then
        DragDrop_BitmapData = Layers.LoadImageAsNewLayer(False, tmpDragDropFile, sTitle, True, , srcX, srcY)
    Else
        DragDrop_BitmapData = LoadFileAsNewImage(tmpDragDropFile, sTitle, False)
    End If
    
    'Be polite and remove the temporary file
    Files.FileDeleteIfExists tmpDragDropFile
    
    If DragDrop_BitmapData Then Message "Image imported successfully "
    
    Exit Function
    
DragDropBMPFailed:
    PDDebug.LogAction "WARNING!  pdClipboardMain.DragDrop_BitmapData() experienced error #" & Err.Number & ": " & Err.Description
    
End Function

'Load dragged text data as a new image and/or layer.  The text will be analyzed for valid file paths or URLs.
Private Function DragDrop_TextData(ByRef OLEDragDrop_DataObject As DataObject, Optional ByVal loadAsLayers As Boolean = False, Optional ByVal suppressErrorMsgs As Boolean = False, Optional ByVal srcX As Long = LONG_MAX, Optional ByVal srcY As Long = LONG_MAX) As Boolean
    
    'On Error GoTo DragDropTextFailed
    
    DragDrop_TextData = False
    
    'Note that PD will validate the data object before passing it, but better safe than sorry
    If (OLEDragDrop_DataObject Is Nothing) Then Exit Function
    If Not OLEDragDrop_DataObject.GetFormat(vbCFText) Then Exit Function
    
    Dim dragText As String
    dragText = OLEDragDrop_DataObject.GetData(vbCFText)
    
    If (LenB(dragText) <> 0) Then
    
        'First, test the text for URL-like indicators
        Dim testURL As String
        testURL = Trim$(dragText)
    
        If Strings.StringsEqualLeft(testURL, "http", True) Or Strings.StringsEqualLeft(testURL, "ftp", True) Then
            
            Message "Image URL found.  Attempting to download..."
            
            'The clipboard URL importer isn't actually specific to the clipboard, so we can freely us it here.
            DragDrop_TextData = ClipboardPaste_WellFormedURL(testURL, loadAsLayers, suppressErrorMsgs)
            
        'If this doesn't look like a URL, see if it's a file path instead
        Else
            
            'Remove leading/trailing quotation marks, if any.
            If (Left$(dragText, 1) = """") Then dragText = Right$(dragText, Len(dragText) - 1)
            If (Right$(dragText, 1) = """") Then dragText = Left$(dragText, Len(dragText) - 1)
            
            Dim targetFile As String
            If Files.FileExists(dragText) Then
                targetFile = dragText
            ElseIf Files.FileExists(Trim$(dragText)) Then
                targetFile = Trim$(dragText)
            End If
            
            'If the text (or a trimmed version of the text) matches a local file, try to load it.
            If (LenB(targetFile) <> 0) Then
                If PDImages.IsImageActive() And loadAsLayers Then
                    DragDrop_TextData = Layers.LoadImageAsNewLayer(False, targetFile, vbNullString, True, True, srcX, srcY)
                Else
                    DragDrop_TextData = LoadFileAsNewImage(targetFile)
                End If
            End If
            
        End If
    
    End If
    
    Exit Function
    
DragDropTextFailed:
    PDDebug.LogAction "WARNING!  pdClipboardMain.DragDrop_TextData() experienced error #" & Err.Number & ": " & Err.Description
    
End Function

'When a Copy or Paste event is initiated, PD doesn't actually copy anything to the clipboard.
' Instead, it "stashes" a copy of the generic, core image data that would be translated to some clipboard format
' (e.g. a layer or standalone DIB or something).  When some other program (or PD itself) requests that data via Paste,
' we then retrieve the "stashed" data and render it into the requested format.  This provides as ton of benefits,
' including better Copy performance (as we only render data in a singular format if/when it's actually needed),
' reduced memory (we don't flood the clipboard with a given format until it's actually required), and cleaner code.
'
'FYI: stashed pixel data uses premultiplied alpha, by design.
Private Sub StashClipboardData(ByVal copyMerged As Boolean)
    
    'The specific data we stash varies according to a few different parameters.
    Dim writeUnAffinedToo As Boolean: writeUnAffinedToo = False
    
    'First, check for an active selection
    If PDImages.GetActiveImage.IsSelectionActive Then
        
        PDDebug.LogAction "Clipboard stash update: generating processed selection area..."
        
        'Fill the temporary DIB with the selected area onlny, with any/all selection processing applied
        PDImages.GetActiveImage.RetrieveProcessedSelection m_ClipboardDIB, True, copyMerged
        
    Else
        
        PDDebug.LogAction "Clipboard stash update: generating composited image and/or layer..."
        
        'If a selection is NOT active, just make a copy of the full layer or image, depending on the merged request
        If copyMerged Then
            PDImages.GetActiveImage.GetCompositedImage m_ClipboardDIB, True
        Else
            If PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True) Then
                writeUnAffinedToo = True
                PDImages.GetActiveImage.GetActiveLayer.GetAffineTransformedDIB m_ClipboardDIB, 0, 0
            Else
                If (m_ClipboardDIB Is Nothing) Then Set m_ClipboardDIB = New pdDIB
                m_ClipboardDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB
            End If
        End If
        
    End If
    
    'Remove any previously stashed file(s).  (This can happen if the user copies, but doesn't paste.)
    If (LenB(m_StashFile) <> 0) Then Files.FileDeleteIfExists m_StashFile
    If (LenB(m_StashFileLayer) <> 0) Then Files.FileDeleteIfExists m_StashFileLayer
    If (LenB(m_StashFileLayerAffineOrig) <> 0) Then Files.FileDeleteIfExists m_StashFileLayerAffineOrig
    If (LenB(m_StashFileVector) <> 0) Then Files.FileDeleteIfExists m_StashFileVector
    
    'Grab a new temporary filename for pixel data
    If (LenB(m_StashFile) = 0) Then m_StashFile = OS.UniqueTempFilename(customExtension:="tmpdib")
    
    PDDebug.LogAction "Writing clipboard stash now..."
    If (Not m_ClipboardDIB Is Nothing) Then
    
        'Dump the temporary DIB to file
        m_ClipboardDIB.WriteToFile m_StashFile, cf_Lz4
        
        'Erase the temporary DIB
        m_ClipboardDIB.EraseDIB
        
        'If the target layer has active affine transforms, save an *untransformed* copy too
        ' (we'll preferentially use it for internal copy+paste ops).
        If writeUnAffinedToo Then
            m_StashFileLayerAffineOrig = OS.UniqueTempFilename(customExtension:="tmpdib")
            PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB.WriteToFile m_StashFileLayerAffineOrig, cf_Lz4
        Else
            Files.FileDeleteIfExists m_StashFileLayerAffineOrig
            m_StashFileLayerAffineOrig = vbNullString
        End If
        
        'Erase any stored layer and vector data
        m_StashedLayer = vbNullString
        m_StashedVector = vbNullString
        
        'If the cut/copy action is an entire source layer, also cache the layer header (and vector data if relevant).
        If (Not PDImages.GetActiveImage.IsSelectionActive) And (Not copyMerged) Then
            
            If (LenB(m_StashFileLayer) = 0) Then m_StashFileLayer = OS.UniqueTempFilename(customExtension:="txt")
            Files.FileSaveAsText PDImages.GetActiveImage.GetActiveLayer.GetLayerHeaderAsXML(), m_StashFileLayer
            
            If PDImages.GetActiveImage.GetActiveLayer.IsLayerVector() Then
                If (LenB(m_StashFileVector) = 0) Then m_StashFileVector = OS.UniqueTempFilename(customExtension:="txt")
                Files.FileSaveAsText PDImages.GetActiveImage.GetActiveLayer.GetVectorDataAsXML(), m_StashFileVector
            End If
        
        'If we're *not* saving layer data, blank out those filenames so we know not to use them later
        Else
            m_StashFileLayer = vbNullString
            m_StashFileVector = vbNullString
        End If
        
        PDDebug.LogAction "Clipboard stashed successfully."
        
    Else
        PDDebug.LogAction "WARNING!  pdClipboardMain.StashClipboardData was passed an empty DIB"
    End If
    
End Sub

'Retrieve previously stashed data.  If this function returns TRUE:
' 1) stashed pixel data has been successfully loaded into m_ClipboardDIB (always guaranteed)
' 2) optionally, stashed layer header (if any) has been successfully loaded into m_stashedLayer
' 3) optionally, stashed vector data (if any) has been successfully loaded into m_stashedVector
Private Function UnstashClipboardData() As Boolean
    
    UnstashClipboardData = False
    
    'Make sure a stash file actually exists
    If (LenB(m_StashFile) <> 0) Then
    
        If Files.FileExists(m_StashFile) Then
        
            'Load the temp file into memory.  This is ALL this function does; actually uploading the data to the clipboard
            ' is handled externally.
            If (m_ClipboardDIB Is Nothing) Then Set m_ClipboardDIB = New pdDIB
            UnstashClipboardData = m_ClipboardDIB.CreateFromFile(m_StashFile)
            
        End If
        
        If Files.FileExists(m_StashFileLayer) Then Files.FileLoadAsString m_StashFileLayer, m_StashedLayer
        If Files.FileExists(m_StashFileVector) Then Files.FileLoadAsString m_StashFileVector, m_StashedVector
        
    End If
    
End Function

'If you know it's safe to destroy PD's clipboard cache, you can do so via this function.
' Note that this action is *not* reversible, so make sure it's what you want/need to do.
Private Sub DestroyStashedData()
    If (LenB(m_StashFile) <> 0) Then Files.FileDeleteIfExists m_StashFile
    If (LenB(m_StashFileLayerAffineOrig) <> 0) Then Files.FileDeleteIfExists m_StashFileLayerAffineOrig
    If (LenB(m_StashFileLayer) <> 0) Then Files.FileDeleteIfExists m_StashFileLayer
    If (LenB(m_StashFileVector) <> 0) Then Files.FileDeleteIfExists m_StashFileVector
End Sub

Friend Function IsPDDataOnClipboard() As Boolean
    IsPDDataOnClipboard = m_Clipboard.IsOurDataOnTheClipboard()
End Function

'Want to render PD's current clipboard stash to the clipboard, WITHOUT delayed rendering?  Call this function.
' (PD does this prior to shutdown, because WM_RENDERALLFORMATS does not play nicely with the order
'  VB unloads objects / destroys windows.)
Friend Sub RenderAllClipboardFormatsManually()
    
    PDDebug.LogAction "pdClipboardMain received notification to render all clipboard formats.  Rendering now..."
    
    'Start by unstashing our previously unsaved clipboard data; this will populate m_ClipboardDIB accordingly
    Dim okayToProceed As Boolean
    okayToProceed = False
    
    If UnstashClipboardData() Then
    
        'Validate the DIB, just to be safe
        If (Not m_ClipboardDIB Is Nothing) Then
            If (m_ClipboardDIB.GetDIBHandle <> 0) And (m_ClipboardDIB.GetDIBWidth <> 0) And (m_ClipboardDIB.GetDIBHeight <> 0) Then
                okayToProceed = True
            End If
        End If
        
    End If
    
    'If we unstashed our data successfully, proceed with clipboard rendering
    If okayToProceed Then
    
        'Open the clipboard
        If m_Clipboard.ClipboardOpen(FormMain.hWnd) Then
            
            'Clear the clipboard
            m_Clipboard.ClearClipboard
            
            'Render all supported formats
            RenderClipboard_PNG
            RenderClipboard_BITMAP
            
            'DIBs have some special considerations.  For details, check out the comments in the
            ' ClipboardCopy() function, but in a nutshell, PD doesn't render CF_DIB images because
            ' various programs handle alpha in varying ways, making it impossible to please everyone.
            ' Instead, as far as standard formats go, PD renders CF_BITMAP and CF_DIBv5 ONLY, with the
            ' assumption that "advanced" image editors can pick up the DIBv5 without trouble, while "basic"
            ' editors (like MS Paint) will take the CF_BITMAP copy.  (This approach also spares the user's
            ' system a good chunk of resources, as we're not copying a crapload of mega-sized images in
            ' varying DIB formats.)
            RenderClipboard_DIB True
            
            'NOTE: regular DIB rendering has been reinstated to workaround issues with Chrome,
            ' which apparently won't grab BITMAP, DIBv5, or PNG formats if available (instead always
            ' defaulting to DIB regardless of its success or failure when retrieving)
            RenderClipboard_DIB False
            
            'Close the clipboard
            m_Clipboard.ClipboardClose
            
            'Release the temporary DIB, if any
            Set m_ClipboardDIB = Nothing
            
        End If
        
    End If
    
    'Because PD has formally placed all of its data on the clipboard, delayed rendering is no longer active, meaning we
    ' can release any remaining clipboard stashes without consequence.
    DestroyStashedData

End Sub

'If we've created a clipboard stash file in the past, remove it
Private Sub m_Clipboard_ClipboardDestroyStashedData()
    
    'Only destroy stashed data if we're not the clipboard owner
    If (Not m_Clipboard.IsOurDataOnTheClipboard) Then DestroyStashedData
    
End Sub

'When the clipboard object tells us to render all remaining clipboard items, WE MUST OPEN THE CLIPBOARD.
' This differs from the behavior of rendering an individual format on-demand.  This is only called when
' the primary PD window is destroyed, because we must upload all supported clipboard data before PD
' disappears.
Private Sub m_Clipboard_ClipboardRenderAllFormats()
    'Because PD handles format rendering manually, we can safely ignore this message at present
End Sub

'When the clipboard object tells us to render a clipboard item on-demand, WE MUST NOT OPEN THE CLIPBOARD.
' Another program has already asserted ownership as part of a Paste operation; we simply plug-in the data
' they need.
Private Sub m_Clipboard_ClipboardRenderFormat(ByVal clipFormatID As Long)

    PDDebug.LogAction "pdClipboardMain received notification to render clipboard format " & m_Clipboard.GetFormatNameFromID(clipFormatID)
    
    Dim okayToProceed As Boolean
    okayToProceed = False
    
    PDDebug.LogAction "Unstashing previously saved clipboard data..."
    
    'Start by unstashing our previously unsaved clipboard data; this will populate m_ClipboardDIB accordingly
    If UnstashClipboardData() Then
    
        'Validate the DIB, just to be safe
        If (Not m_ClipboardDIB Is Nothing) Then
            okayToProceed = ((m_ClipboardDIB.GetDIBHandle <> 0) And (m_ClipboardDIB.GetDIBWidth <> 0) And (m_ClipboardDIB.GetDIBHeight <> 0))
        End If
        
    End If
    
    'If the DIB passed validation, proceed with rendering
    If okayToProceed Then
    
        Select Case clipFormatID
            
            Case CF_BITMAP
                RenderClipboard_BITMAP
            
            Case m_Clipboard.AddClipboardFormat("PNG")
                RenderClipboard_PNG
            
            Case CF_DIB, CF_DIBV5
                RenderClipboard_DIB (clipFormatID = CF_DIBV5)
            
            Case Else
                PDDebug.LogAction "WARNING!  A clipboard render request was received for an unhandled format: " & clipFormatID
                
        End Select
        
    End If
    
    'Release any temporary DIBs created by the request
    Set m_ClipboardDIB = Nothing

End Sub

'Place a PNG on the clipboard.  This function assumes that the caller has already unstashed previously
' saved clipboard data, and thoroughly validated it.  If you haven't done that, this function will fail.
' NOTE: this function does not open, clear, or close the clipboard, by design; it is up to the caller
' to handle that.
Private Sub RenderClipboard_PNG()
    
    On Error GoTo RenderClipboardPNG_Failure
    If (m_ClipboardDIB Is Nothing) Then
        PDDebug.LogAction "WARNING!  RenderClipboard_BITMAP() failed; DIB empty"
        Exit Sub
    End If
    
    'The stashed DIB will likely have premultiplied alpha.  Un-premultiply it now.
    If m_ClipboardDIB.GetAlphaPremultiplication() Then m_ClipboardDIB.SetAlphaPremultiplication False
    
    PDDebug.LogAction "Clipboard copy update: retrieving PNG format ID..."
    
    'Most systems will already have PNG available as a setting; the AddFormat function will detect this
    Dim pngID As Long
    pngID = m_Clipboard.AddClipboardFormat("PNG")
    
    PDDebug.LogAction "Clipboard copy update: using pdPNG to produce PNG stream..."
    
    'Use our awesome homebrew PNG encoder to quickly produce a PNG stream from the source image data
    Dim pngArray() As Byte, sizePngStream As Long
    
    Dim cPNG As pdPNG
    Set cPNG = New pdPNG
    sizePngStream = cPNG.SavePNG_ToMemory(dstArray:=pngArray, srcImage:=m_ClipboardDIB, srcPDImage:=Nothing, outColorType:=png_TruecolorAlpha, outBitsPerChannel:=8, cmpLevel:=1, useFilterStrategy:=png_FilterNone)
    
    If (sizePngStream <> 0) Then
        PDDebug.LogAction "Clipboard copy update: uploading PNG to clipboard..."
        If m_Clipboard.SetClipboardBinaryData(pngID, pngArray) Then PDDebug.LogAction "Clipboard copy update: PNG clipboard operation complete."
    Else
        PDDebug.LogAction "WARNING!  During ClipboardCopy, pdPNG failed to produce a usable PNG stream.  PNG format abandoned."
    End If
    
RenderClipboardPNG_Failure:
    
End Sub

'Place an old-fashioned hBitmap on the clipboard.  This function assumes that the caller has already unstashed previously saved
' clipboard data and thoroughly validated it.  If you haven't done that, this function will fail.
' NOTE: this function does not open, clear, or close the clipboard, by design; it is up to the caller to handle that.
Private Sub RenderClipboard_BITMAP()
    
    On Error GoTo RenderClipboardBITMAP_Failure
    If (m_ClipboardDIB Is Nothing) Then
        PDDebug.LogAction "WARNING!  RenderClipboard_BITMAP() failed; DIB empty"
        Exit Sub
    End If
    
    'Get a handle to the current desktop, and create a compatible clipboard device context in it
    Dim desktopHWnd As Long, tmpDC As Long
    desktopHWnd = GetDesktopWindow()
    tmpDC = GetDC(desktopHWnd)
    
    If (tmpDC <> 0) Then
        
        PDDebug.LogAction "Clipboard copy update: creating DDB copy of image..."
        
        'Create a DDB compatible with the current desktop. This will receive the actual pixel data of the current DIB,
        ' downsampled to 24-bpp.
        Dim clipboardDC As Long, clipboardBMP As Long
        clipboardDC = GDI.GetMemoryDC(tmpDC)
        
        If (clipboardDC <> 0) Then
        
            clipboardBMP = CreateCompatibleBitmap(tmpDC, m_ClipboardDIB.GetDIBWidth, Abs(m_ClipboardDIB.GetDIBHeight))
            If (clipboardBMP <> 0) Then
                
                PDDebug.LogAction "Clipboard copy update: compositing 24-bpp copy for DDB format..."
                
                'Select the DDB into our temporary compatible DC so we can draw to it with GDI commands
                Dim clipboardOldBMP As Long
                clipboardOldBMP = SelectObject(clipboardDC, clipboardBMP)
                
                '24-bit images can be copied as-is
                If (m_ClipboardDIB.GetDIBColorDepth = 24) Then
                    GDI.BitBltWrapper clipboardDC, 0, 0, m_ClipboardDIB.GetDIBWidth, m_ClipboardDIB.GetDIBHeight, m_ClipboardDIB.GetDIBDC, 0, 0, vbSrcCopy
                    
                '32-bit images must be composited against a white background first
                Else
                    
                    'Create a 24-bpp copy of the image, then render it onto our DDB
                    Dim tmpDIB As pdDIB
                    Set tmpDIB = New pdDIB
                    tmpDIB.CreateFromExistingDIB m_ClipboardDIB
                    tmpDIB.ConvertTo24bpp
                    
                    GDI.BitBltWrapper clipboardDC, 0, 0, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
                    
                    Set tmpDIB = Nothing
                    
                End If
                
                PDDebug.LogAction "Clipboard copy update: uploading DDB to clipboard..."
                
                'Remove the clipboard bitmap from the temporary DC
                SelectObject clipboardDC, clipboardOldBMP
                
                'Copy the DDB into the clipboard; the clipboard now assumes ownership over the DDB
                m_Clipboard.SetClipboardMemoryHandle CF_BITMAP, clipboardBMP, False
                
                'There is some debate over whether or not we should delete our copy of the DDB handle
                ' (see: http://stackoverflow.com/questions/32086618/who-releases-handle-in-setclipboarddatacf-bitmap-hbitmap)
                ' but since deleting it causes no problems, I figure that's safest.
                DeleteObject clipboardBMP
                GDI.FreeMemoryDC clipboardDC
                
            End If
            
        Else
            PDDebug.LogAction "WARNING!  Prior to rendering BITMAP clipboard data, a compatible DC could not be created.  Copy abandoned."
        End If
        
        PDDebug.LogAction "Clipboard copy update: freeing temporary clipboard resources..."
        ReleaseDC desktopHWnd, tmpDC
        
    End If
    
    PDDebug.LogAction "Clipboard copy update: DDB clipboard operation complete."
    
RenderClipboardBITMAP_Failure:

End Sub

'Place a DIB (or optionally, DIBv5) on the clipboard.  This function assumes that the caller has already unstashed previously
' saved clipboard data and thoroughly validated it.  If you haven't done that, this function will fail.
' NOTE: this function does not open, clear, or close the clipboard, by design; it is up to the caller to handle that.
Private Sub RenderClipboard_DIB(Optional ByVal useV5Header As Boolean = False)
    
    On Error GoTo RenderClipboardDIB_Failure
    If (m_ClipboardDIB Is Nothing) Then
        PDDebug.LogAction "WARNING!  RenderClipboard_DIB() failed; DIB empty"
        Exit Sub
    End If
    
    If useV5Header Then
        PDDebug.LogAction "Clipboard copy update: allocating global memory for DIBv5 object..."
    Else
        PDDebug.LogAction "Clipboard copy update: allocating global memory for DIB object..."
    End If
    
    'DIBs should be unpremultiplied prior to copying; note that some esoteric software
    ' (*cough* XNView *cough*) wants premultiplied alpha, but the general consensus seems to be
    ' "use unpremultiplied", so that's what we do too.
    '
    'Note that alpha state is only relevant for DIBv5 DIBs; for compatibility reasons, PD always
    ' composites plain DIBs against a white backdrop, since alpha channel compatibility is so
    ' variable between apps.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    If useV5Header Then
        If m_ClipboardDIB.GetAlphaPremultiplication Then m_ClipboardDIB.SetAlphaPremultiplication False
    Else
        tmpDIB.CreateFromExistingDIB m_ClipboardDIB
        tmpDIB.CompositeBackgroundColor 255, 255, 255
    End If
    
    'Figure out how much size is required for the global allocation.  This is just (size_of_header + size_of_pixels).
    Dim headerSize As Long
    Dim bmpHeader As BITMAPINFOHEADER, bmpV5Header As BITMAPV5HEADER
    If useV5Header Then
        headerSize = LenB(bmpV5Header)
    Else
        headerSize = LenB(bmpHeader)
    End If
    
    Dim dibPointer As Long, dibSize As Long
    If useV5Header Then
        m_ClipboardDIB.RetrieveDIBPointerAndSize dibPointer, dibSize
    Else
        tmpDIB.RetrieveDIBPointerAndSize dibPointer, dibSize
    End If
    
    Dim memSize As Long
    memSize = headerSize + dibSize
    
    'Allocate sufficient global memory for the header and pixels
    Dim hGlobal As Long, hGlobalPtr As Long
    hGlobal = m_Clipboard.GetGlobalMemoryHandle(memSize)
    
    If (hGlobal <> 0) Then
        hGlobalPtr = m_Clipboard.GetPointerFromGlobalMemory(hGlobal)
        If (hGlobalPtr <> 0) Then
            
            PDDebug.LogAction "Clipboard copy update: copying DIB and header into global memory..."
            
            'Populate the relevant DIB header
            If useV5Header Then
                With bmpV5Header
                    .biSize = LenB(bmpV5Header)
                    .biWidth = m_ClipboardDIB.GetDIBWidth
                    .biHeight = m_ClipboardDIB.GetDIBHeight
                    .biPlanes = 1
                    .biBitCount = m_ClipboardDIB.GetDIBColorDepth
                    .biXPelsPerMeter = PDMath.ConvertDPIToPels(m_ClipboardDIB.GetDPI)
                    .biYPelsPerMeter = .biXPelsPerMeter
                    
                    'Even though we don't specify BI_BITFIELDS as the compression scheme, some software will disregard DIBv5 alpha
                    ' if no alpha mask is provided.  As such, let's provide default mask values, "just in case".
                    .biAlphaMask = RGBA_ALPHA_MASK
                    .biBlueMask = RGBA_BLUE_MASK
                    .biGreenMask = RGBA_GREEN_MASK
                    .biRedMask = RGBA_RED_MASK
                    
                End With
            Else
                With bmpHeader
                    .biSize = LenB(bmpHeader)
                    .biWidth = m_ClipboardDIB.GetDIBWidth
                    .biHeight = m_ClipboardDIB.GetDIBHeight
                    .biPlanes = 1
                    .biBitCount = m_ClipboardDIB.GetDIBColorDepth
                    .biXPelsPerMeter = PDMath.ConvertDPIToPels(m_ClipboardDIB.GetDPI)
                    .biYPelsPerMeter = .biXPelsPerMeter
                End With
            End If
            
            'Copy the header into global memory
            If useV5Header Then
                CopyMemoryStrict hGlobalPtr, VarPtr(bmpV5Header), LenB(bmpV5Header)
            Else
                CopyMemoryStrict hGlobalPtr, VarPtr(bmpHeader), LenB(bmpHeader)
            End If
            
            'Next, we want to copy the actual pixel bits into place.  PD stores DIBs as bottom-up, but this can cause problems when
            ' programs expect a top-down DIB.  To prevent trouble, we'll reverse the order of lines.
            Dim y As Long, yFinal As Long, dibStride As Long
            yFinal = m_ClipboardDIB.GetDIBHeight - 1
            dibStride = m_ClipboardDIB.GetDIBStride
            
            For y = 0 To yFinal
                CopyMemoryStrict hGlobalPtr + headerSize + (y * dibStride), dibPointer + (yFinal - y) * dibStride, dibStride
            Next y
            
            'Unlock the memory, then assign ownership to the clipboard
            m_Clipboard.FinishedWithGlobalMemoryPtr hGlobal
            
            PDDebug.LogAction "Clipboard copy update: transferring ownership to the clipboard..."
            
            If useV5Header Then
                m_Clipboard.SetClipboardMemoryHandle CF_DIBV5, hGlobal
            Else
                m_Clipboard.SetClipboardMemoryHandle CF_DIB, hGlobal
            End If
            
        Else
            PDDebug.LogAction "WARNING!  pdClipboardMain.RenderClipboard_DIB couldn't lock its global memory handle."
        End If
        
    Else
        PDDebug.LogAction "WARNING!  pdClipboardMain.RenderClipboard_DIB couldn't allocate enough global memory for this operation."
    End If
    
    PDDebug.LogAction "Clipboard copy update: DIB clipboard operation complete."
    
RenderClipboardDIB_Failure:

End Sub

Private Sub Class_Initialize()

    Set m_Clipboard = New pdClipboard
    
    'While it's not exactly time-consuming to retrieve custom clipboard formats, we may as well cache a few common ones.
    ' (This is made cumbersome by a lack of an unsigned Int format; as such, we have to jiggle the values around to prevent
    '  VB from throwing "out of range" errors inside the IDE - CF_FileContents is declared as a 16-bit integer, because it's
    '  passed to the internal VB function DataObject.GetFormat().)
    CF_FileContents = m_Clipboard.AddClipboardFormat("FileContents") And &H7FFF& Or &H8000
    CF_FileGroupDescriptorW = m_Clipboard.AddClipboardFormat("FileGroupDescriptorW") And &H7FFF& Or &H8000
    
End Sub

Private Sub Class_Terminate()
    
    Set m_Clipboard = Nothing
    
    'Perform a failsafe check against un-destroyed stash data
    If (LenB(m_StashFile) <> 0) Then Files.FileDeleteIfExists m_StashFile
    If (LenB(m_StashFileLayerAffineOrig) <> 0) Then Files.FileDeleteIfExists m_StashFileLayerAffineOrig
    If (LenB(m_StashFileLayer) <> 0) Then Files.FileDeleteIfExists m_StashFileLayer
    If (LenB(m_StashFileVector) <> 0) Then Files.FileDeleteIfExists m_StashFileVector
    
End Sub
