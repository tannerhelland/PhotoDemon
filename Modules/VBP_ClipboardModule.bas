Attribute VB_Name = "Clipboard_Handler"
'***************************************************************************
'Clipboard Interface
'Copyright 2001-2015 by Tanner Helland
'Created: 15/April/01
'Last updated: 15/November/15
'Last update: finalize massive overhaul to integrate new pdClipboard class during PASTE operations.  Cut/Copy still need work.
'
'Module for handling all Windows clipboard routines.  Most clipboard APIs are not located here, but in the separate pdClipboard
' object, which includes a ton of specialized helper functions.  Look there for details.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Some specialized structs are required for parsing various clipboard bitmap formats
Private Type BITMAPFILEHEADER
    Type As Integer
    Size As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffBits As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
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

Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal dwWidth As Long, ByVal dwHeight As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, ByVal lpvBits As Long, ByVal lpbmi As Long, ByVal fuColorUse As Long) As Long

'If the clipboard is currently open, this module will cache certain values.  External functions (like the image load routines) can query these
' for additional information on how to handle the image.  Clipboard images have a ton of caveats that normal images do not; hence these trackers.
Private m_IsClipboardOpen As Boolean

'If external functions want to know more about the current clipboard operation, they can retrieve a copy of this struct.
Public Type PD_CLIPBOARD_INFO
    pdci_CurrentFormat As PredefinedClipboardFormatConstants
    pdci_OriginalFormat As PredefinedClipboardFormatConstants
    pdci_DIBv5AlphaMask As Long
End Type

Private m_ClipboardInfo As PD_CLIPBOARD_INFO

'IMPORTANT NOTE: at present, this value is only updated during paste steps.  I still need to convert cut/copy operations to use pdClipboard.
Public Function IsClipboardOpen() As Boolean
    IsClipboardOpen = m_IsClipboardOpen
End Function

Public Function GetClipboardInfo() As PD_CLIPBOARD_INFO
    GetClipboardInfo = m_ClipboardInfo
End Function

'Copy the current selection (or entire layer, if no selection is active) to the clipboard, then erase the selected area
' (or entire layer, if no selection is active).
Public Sub ClipboardCut(ByVal cutMerged As Boolean)

    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Check for an active selection
    If pdImages(g_CurrentImage).selectionActive Then
    
        'Fill the temporary DIB with the selection
        pdImages(g_CurrentImage).retrieveProcessedSelection tmpDIB, False, cutMerged
        
    Else
        
        'If a selection is NOT active, just make a copy of the full layer or image, depending on the merged request
        If cutMerged Then
            pdImages(g_CurrentImage).getCompositedImage tmpDIB, False
        Else
            tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).getActiveLayer.layerDIB
            
            'Layers are always premultiplied, so we must unpremultiply it now if 32bpp
            If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.setAlphaPremultiplication False
            
        End If
        
    End If
    
    'Copy the temporary DIB to the clipboard, then erase it
    DIB_Handler.copyDIBToClipboard tmpDIB
    Set tmpDIB = Nothing
    
    'Now, we have the added step of erasing the selected area from the screen.  "Cut merged" requires us to delete the selected
    ' region from all visible layers, so vary the loop bounds accordingly.
    Dim startLayer As Long, endLayer As Long
    
    If cutMerged Then
        startLayer = 0
        endLayer = pdImages(g_CurrentImage).getNumOfLayers - 1
    Else
        startLayer = pdImages(g_CurrentImage).getActiveLayerIndex
        endLayer = pdImages(g_CurrentImage).getActiveLayerIndex
    End If
    
    Dim i As Long
    For i = startLayer To endLayer
        
        'For "cut merged", ignore transparent layers
        If cutMerged Then
        
            If pdImages(g_CurrentImage).getLayerByIndex(i).getLayerVisibility Then
                
                'If a selection is active, erase the selected area.  Otherwise, wipe the whole layer.
                If pdImages(g_CurrentImage).selectionActive Then
                    pdImages(g_CurrentImage).eraseProcessedSelection i
                Else
                    Layer_Handler.eraseLayerByIndex i
                End If
                
            End If
        
        'For "cut from layer", erase the selection regardless of layer visibility
        Else
        
            'If a selection is active, erase the selected area.  Otherwise, delete the given layer.
            If pdImages(g_CurrentImage).selectionActive Then
                pdImages(g_CurrentImage).eraseProcessedSelection i
            Else
                Layer_Handler.deleteLayer i
            End If
            
        End If
        
    Next i
    
    'Redraw the active viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Copy the current layer (or composite image, if copyMerged is true) to the clipboard.
Public Sub ClipboardCopy(ByVal copyMerged As Boolean)
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Check for an active selection
    If pdImages(g_CurrentImage).selectionActive Then
    
        'Fill the temporary DIB with the selection
        pdImages(g_CurrentImage).retrieveProcessedSelection tmpDIB, False, copyMerged
        
    Else
    
        'If a selection is NOT active, just make a copy of the full layer or image, depending on the merged request
        If copyMerged Then
            pdImages(g_CurrentImage).getCompositedImage tmpDIB, False
        Else
            tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).getActiveLayer.layerDIB
            
            'Layers are always premultiplied, so we must unpremultiply it now if 32bpp
            If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.setAlphaPremultiplication False
            
        End If
        
    End If
    
    'Copy the temporary DIB to the clipboard, then erase it
    DIB_Handler.copyDIBToClipboard tmpDIB
    tmpDIB.eraseDIB
    
End Sub

'Empty the clipboard
Public Sub ClipboardEmpty()
    Clipboard.Clear
End Sub

'Paste an image (e.g. create new image data based on whatever is in the clipboard).
' The parameter "srcIsMeantAsLayer" controls whether the clipboard data is loaded as a new image, or as a new layer in an existing image.
Public Sub ClipboardPaste(ByVal srcIsMeantAsLayer As Boolean)
    
    Dim pasteWasSuccessful As Boolean
    pasteWasSuccessful = False
    
    'Prep a bunch of generic clipboard handling variables
    Dim tmpClipboardFile As String, tmpDownloadFile As String
    Dim sFile(0) As String
    Dim sTitle As String, sFilename As String
    
    'Also, note that all file interactions occur through pdFSO, so we can support Unicode filenames/paths.
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'Attempt to open the clipboard
    Dim clpObject As pdClipboard
    Set clpObject = New pdClipboard
    If clpObject.ClipboardOpen(FormMain.hWnd) Then
        
        'Mark the clipboard as open; external functions can query this value
        m_IsClipboardOpen = True
        
        'When debugging, it's nice to know what clipboard formats the OS reports prior to actually retrieving them.
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Clipboard reports the following formats: " & clpObject.GetListOfAvailableFormatNames()
        #End If
        
        'PNGs on the clipboard get preferential treatment, as they preserve transparency data - so check for them first.
        If clpObject.DoesClipboardHaveFormatName("PNG") Then
            pasteWasSuccessful = ClipboardPaste_CustomImageFormat(clpObject, "PNG", srcIsMeantAsLayer, "png")
        End If
        
        'If we couldn't find PNG data (or something went horribly wrong during that step), look for an HTML fragment next.
        ' Images copied from web browsers typically create an HTML fragment, which should have a direct link to the copied image.
        '  Downloading the image manually lets us maintain things like ICC profiles and the image's original filename.
        If clpObject.DoesClipboardHaveHTML() And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = ClipboardPaste_HTML(clpObject, srcIsMeantAsLayer)
        End If
        
        'JPEGs are another possibility.  We prefer them less than PNG or direct download (because there's no guarantee that the
        ' damn browser didn't re-encode them, but they're better than bitmaps or DIBs because they may retain metadata and
        ' color profiles, so test for JPEG next.  (Also, note that certain versions of Microsoft Office use "JFIF" as the identifier,
        ' for reasons known only to them...)
        If clpObject.DoesClipboardHaveFormatName("JPEG") And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = ClipboardPaste_CustomImageFormat(clpObject, "JPEG", srcIsMeantAsLayer, "jpg")
        End If
        
        If clpObject.DoesClipboardHaveFormatName("JPG") And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = ClipboardPaste_CustomImageFormat(clpObject, "JPG", srcIsMeantAsLayer, "jpg")
        End If
        
        If clpObject.DoesClipboardHaveFormatName("JFIF") And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = ClipboardPaste_CustomImageFormat(clpObject, "JFIF", srcIsMeantAsLayer, "jpg")
        End If
        
        'Next, see if the clipboard contains a generic file list.  If it does, try to load each file in turn.
        If clpObject.DoesClipboardHaveFiles() And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = ClipboardPaste_ListOfFiles(clpObject, srcIsMeantAsLayer)
        End If
        
        'Next, look for plaintext.  This could be a URL, or maybe a text representation of a filepath.
        ' (Also, note that we only have to search for one text format, because the OS auto-converts between text formats for free.)
        If clpObject.DoesClipboardHaveText() And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = ClipboardPaste_TextSource(clpObject, srcIsMeantAsLayer)
        End If
        
        'Last up are DIBs and bitmaps.  Once again, the OS auto-converts between bitmap and DIB formats, and if it all possible,
        ' we prefer DIBv5 as it actually supports alpha data.
        If clpObject.DoesClipboardHaveBitmapImage() And (Not pasteWasSuccessful) Then
            pasteWasSuccessful = ClipboardPaste_BitmapImage(clpObject, srcIsMeantAsLayer)
        End If
        
        'Regardless of success or failure, make sure to close the clipboard now that we're done with it.
        clpObject.ClipboardClose
        
        'Mark the clipboard as closed
        m_IsClipboardOpen = False
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  Couldn't open the clipboard; is it possible another program has locked it?"
        #End If
    End If
        
    'If a paste operation was successful, switch the current tool to the layer move/resize tool, which is most likely needed after a
    ' new layer has been pasted.
    If pasteWasSuccessful Then
        If srcIsMeantAsLayer Then toolbar_Toolbox.selectNewTool NAV_MOVE
    Else
        PDMsgBox "The clipboard is empty or it does not contain a valid picture format.  Please copy a valid image onto the clipboard and try again.", vbExclamation + vbOKOnly + vbApplicationModal, "Windows Clipboard Error"
    End If
    
End Sub

'If the clipboard contains custom-format image data (most commonly PNG or JPEG), you can call this function to initiate a "paste" command
' using the custom image data as a source.  The parameter "srcIsMeantAsLayer" controls whether the clipboard data is loaded as a new image,
' or as a new layer in an existing image.
'
'RETURNS: TRUE if successful; FALSE otherwise.
Private Function ClipboardPaste_CustomImageFormat(ByRef clpObject As pdClipboard, ByVal clipboardFormatName As String, ByVal srcIsMeantAsLayer As Boolean, Optional ByVal tmpFileExtension As String = "tmp") As Boolean
    
    'Unfortunately, a lot of things can go wrong when pasting custom image data, so we assume failure by default.
    ClipboardPaste_CustomImageFormat = False
    
    'All paste operations use a few consistent variables
    
    'Raw retrieval storage variables
    Dim clipFormatID As Long, rawClipboardData() As Byte
    
    'Temporary file for storing the clipboard data.  (This lets us use PD's central image load function.)
    Dim tmpClipboardFile As String
    
    'Additional file information variables, which we pass to the central load function to let it know that this is only a temp file,
    ' and it should use these hint values instead of assuming normal image load behavior.
    Dim sFile() As String, sTitle As String, sFilename As String
    ReDim sFile(0) As String
    
    'pdFSO is used to ensure Unicode subfolder compatibility
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'Verify that the requested data is actually available.  (Hopefully the caller already checked this, but you never know...)
    If clpObject.DoesClipboardHaveFormatName(clipboardFormatName) Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ClipboardPaste_CustomImageFormat() will now attempt to load " & clipboardFormatName & " from the clipboard..."
        #End If
        
        'Because custom-format image data can be registered by many programs, retrieve this image format's unique ID now.
        clipFormatID = clpObject.GetFormatIDFromName(clipboardFormatName)
        m_ClipboardInfo.pdci_CurrentFormat = clipFormatID
        m_ClipboardInfo.pdci_OriginalFormat = clipFormatID
        
        'Pull the data into a local array
        If clpObject.GetClipboardBinaryData(clipFormatID, rawClipboardData) Then
            
            'Dump the data out to file
            tmpClipboardFile = g_UserPreferences.GetTempPath & "PDClipboard." & tmpFileExtension
            If cFile.SaveByteArrayToFile(rawClipboardData, tmpClipboardFile) Then
                
                'We no longer need our local copy of the clipboard data
                Erase rawClipboardData
                
                'We can now use the standard image load routine to import the temporary file.  Because we don't want the
                ' load function to use the temporary file name as the image name, we manually supply a filename to suggest
                ' if the user eventually tries to save the file.
                sFile(0) = tmpClipboardFile
                sTitle = g_Language.TranslateMessage("Clipboard Image")
                sFilename = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
                
                'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                If srcIsMeantAsLayer Then
                    Layer_Handler.loadImageAsNewLayer False, sFile(0), sTitle, True
                Else
                    LoadFileAsNewImage sFile, False, sTitle, sFilename
                End If
                    
                'Be polite and remove the temporary file
                cFile.KillFile tmpClipboardFile
                
                'If we made it all the way here, the load was (probably?) successful
                Message "Clipboard data imported successfully "
                ClipboardPaste_CustomImageFormat = True
                
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  Clipboard image data (probably PNG) could not be written to a temp file."
                #End If
            End If
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  Clipboard.GetBinaryData failed on custom image data (probably PNG).  Special paste action abandoned."
            #End If
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  ClipboardPaste_CustomImageFormat was called, but the requested data doesn't exist on the clipboard."
        #End If
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
Private Function ClipboardPaste_HTML(ByRef clpObject As pdClipboard, ByVal srcIsMeantAsLayer As Boolean) As Boolean
    
    'Unfortunately, a lot of things can go wrong when pasting custom image data, so we assume failure by default.
    ClipboardPaste_HTML = False
    
    'Verify that the requested data is actually available.  (Hopefully the caller already checked this, but you never know...)
    If clpObject.DoesClipboardHaveHTML() Then
        
        'HTML handling requires no special behavior on the part of external load functions, so we mark the module-level tracker as blank
        m_ClipboardInfo.pdci_CurrentFormat = 0
        m_ClipboardInfo.pdci_OriginalFormat = 0
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ClipboardPaste_HTML() will now attempt to find valid image HTML on the clipboard..."
        #End If
        
        'Pull the HTML data into a local string
        Dim htmlString As String
        If clpObject.GetClipboardHTML(htmlString) Then
            
            'Look for an image reference within the HTML snippet
            If InStr(1, UCase$(htmlString), "<IMG ", vbBinaryCompare) > 0 Then
            
                'Retrieve the full image path, which will be between the first set of quotation marks following the
                ' "<img src=" statement in the HTML snippet.
                Dim vbQuoteMark As String
                vbQuoteMark = """"
                
                'Parse out the URL between the img src quotes
                Dim urlStart As Long, urlEnd As Long
                urlStart = InStr(1, UCase$(htmlString), "<IMG ", vbBinaryCompare)
                If urlStart > 0 Then urlStart = InStr(urlStart, UCase$(htmlString), "SRC=", vbBinaryCompare)
                If urlStart > 0 Then urlStart = InStr(urlStart, htmlString, vbQuoteMark, vbBinaryCompare) + 1
                
                'The magic number 6 below is calculated as the length of (src="), + 1 to advance to the
                ' character immediately following the quotation mark.
                If urlStart > 0 Then urlEnd = InStr(urlStart + 6, htmlString, vbQuoteMark, vbBinaryCompare)
                
                'As a failsafe, make sure a valid URL was actually found
                If (urlStart > 0) And (urlEnd > 0) Then
                    
                    ClipboardPaste_HTML = ClipboardPaste_WellFormedURL(Mid$(htmlString, urlStart, urlEnd - urlStart), srcIsMeantAsLayer)
                
                'An image tag was found, but a parsing error occurred when trying to strip out the source URL.  This is okay;
                ' exit immediately without raising any errors.
                Else
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "Clipboard.GetClipboardHTML was successful and an image URL was located, but a parsing error occurred."
                    #End If
                End If
                
            'No image tag found, which is fine; exit immediately without raising any errors.
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Clipboard.GetClipboardHTML was successful, but no image URL found."
                #End If
            End If
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  Clipboard.GetClipboardHTML failed.  Special paste action abandoned."
            #End If
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  ClipboardPaste_HTML was called, but HTML data doesn't exist on the clipboard."
        #End If
    End If
    
End Function

'If one or more files exist on the clipboard, attempt to paste them all.
Private Function ClipboardPaste_ListOfFiles(ByRef clpObject As pdClipboard, ByVal srcIsMeantAsLayer As Boolean)
    
    'Make sure files actually exist on the clipboard
    If clpObject.DoesClipboardHaveFiles() Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ClipboardPaste_ListOfFiles() will now attempt to load one or more files from the clipboard..."
        #End If
        
        'File lists require no special behavior on the part of external load functions, so we mark the module-level tracker as blank
        m_ClipboardInfo.pdci_CurrentFormat = CF_HDROP
        m_ClipboardInfo.pdci_OriginalFormat = CF_HDROP
        
        Dim listOfFiles() As String, numOfFiles As Long
        If clpObject.GetFileList(listOfFiles, numOfFiles) Then
            
            'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
            If srcIsMeantAsLayer Then
                Dim i As Long
                For i = 0 To UBound(listOfFiles)
                    If Len(listOfFiles(i)) <> 0 Then Layer_Handler.loadImageAsNewLayer False, listOfFiles(i), , True
                Next i
            Else
                LoadFileAsNewImage listOfFiles
            End If
            
            ClipboardPaste_ListOfFiles = True
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  ClipboardPaste_ListOfFiles couldn't retrieve a valid file list from pdClipboard."
            #End If
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  ClipboardPaste_ListOfFiles was called, but no file paths exist on the clipboard."
        #End If
    End If
        
End Function

'If the clipboard contains text, try to find an image path or URL that we can use.
Private Function ClipboardPaste_TextSource(ByRef clpObject As pdClipboard, ByVal srcIsMeantAsLayer As Boolean)
    
    ClipboardPaste_TextSource = False
    
    'Make sure text actually exists on the clipboard
    If clpObject.DoesClipboardHaveText() Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ClipboardPaste_TextSource() will now parse clipboard text, looking for image sources..."
        #End If
        
        'Text requires no special behavior on the part of external load functions, so we mark the module-level tracker as blank
        m_ClipboardInfo.pdci_CurrentFormat = 0
        m_ClipboardInfo.pdci_OriginalFormat = 0
        
        Dim clipText As String
        If clpObject.GetClipboardText(clipText) Then
            
            'First, test the text for URL-like indicators
            Dim testURL As String
            testURL = Trim$(clipText)
        
            If (StrComp(UCase$(Left$(testURL, 4)), "HTTP", vbBinaryCompare) = 0) Or (StrComp(UCase$(Left$(testURL, 3)), "FTP", vbBinaryCompare) = 0) Then
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Possible image URL found on clipboard.  Attempting to retrieve data..."
                #End If
                
                Message "Image URL found on clipboard.  Attempting to download..."
                
                ClipboardPaste_TextSource = ClipboardPaste_WellFormedURL(testURL, srcIsMeantAsLayer)
                
            'If this doesn't look like a URL, see if it's a file path instead
            Else
                
                Dim targetFile As String
                targetFile = ""
                
                Dim cFile As pdFSO
                Set cFile = New pdFSO
                If cFile.FileExist(clipText) Then
                    targetFile = clipText
                ElseIf cFile.FileExist(Trim$(clipText)) Then
                    targetFile = Trim$(clipText)
                End If
                
                'If the text (or a trimmed version of the text) matches a local file, try to load it.
                If Len(targetFile) <> 0 Then
                
                    If srcIsMeantAsLayer Then
                        Layer_Handler.loadImageAsNewLayer False, targetFile, , True
                    Else
                        Dim tmpFiles() As String
                        ReDim tmpFiles(0) As String
                        tmpFiles(0) = targetFile
                        LoadFileAsNewImage tmpFiles
                    End If
                    
                    ClipboardPaste_TextSource = True
                
                End If
                
            End If
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  ClipboardPaste_TextSource couldn't retrieve actual text from pdClipboard."
            #End If
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  ClipboardPaste_TextSource was called, but no text exists on the clipboard."
        #End If
    End If
        
End Function

'Helper function that SHOULD NOT BE CALLED DIRECTLY.  Only other ClipboardPaste_* variants are able to safely use this function.
' Returns: TRUE if an image was successfully downloaded and loaded to the pdImages collection.  FALSE if failure occurred.
Private Function ClipboardPaste_WellFormedURL(ByVal srcURL As String, ByVal srcIsMeantAsLayer As Boolean) As Boolean
    
    'This function assumes the source URL is both absolute and well-formed
    Message "Image URL found on clipboard.  Attempting to download..."
                    
    Dim tmpDownloadFile As String
    tmpDownloadFile = FormInternetImport.downloadURLToTempFile(srcURL)
    
    'pdFSO is used to ensure Unicode filename compatibility
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'If the download was successful, we can now use the standard image load routine to import the temporary file
    If cFile.FileLenW(tmpDownloadFile) <> 0 Then
        
        'Additional file information variables, which we pass to the central load function to let it know that this is only a temp file,
        ' and it should use these hint values instead of assuming normal image load behavior.
        Dim sFile() As String
        ReDim sFile(0) As String
        sFile(0) = tmpDownloadFile
                
        Dim tmpFilename As String
        tmpFilename = cFile.GetFilename(tmpDownloadFile, True)
        
        'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
        If srcIsMeantAsLayer Then
            Layer_Handler.loadImageAsNewLayer False, sFile(0), True
        Else
            LoadFileAsNewImage sFile, False, tmpFilename, tmpFilename, , , , , , True
        End If
        
        'Delete the temporary file
        cFile.KillFile tmpDownloadFile
        
        Message "Clipboard data imported successfully "
            
        'Check for load failure.  If the most recent pdImages object is inactive, it's a safe assumption that
        ' the load operation failed.  (This isn't foolproof, especially if the user loads a ton of images,
        ' and subsequently unloads images in an arbitrary order - but given the rarity of this situation, I'm content
        ' to use this technique for predicting failure.)
        If g_CurrentImage <= UBound(pdImages) Then
            If Not pdImages(g_CurrentImage) Is Nothing Then
                If pdImages(g_CurrentImage).IsActive Then
                    ClipboardPaste_WellFormedURL = True
                End If
            End If
        End If
    
    'If the download failed, let the user know that hey, at least we tried.
    Else
        Message "Image download failed.  Please copy a valid image URL to the clipboard and try again."
    End If
    
End Function

'If the clipboard contains bitmap-format image data (or by extension, DIB or DIBv5), you can call this function to initiate a "paste" command.
' The function will automatically determine the best format for pasting.  The parameter "srcIsMeantAsLayer" controls whether the clipboard data
' is loaded as a new image, or as a new layer in an existing image.
'
'RETURNS: TRUE if successful; FALSE otherwise.
Private Function ClipboardPaste_BitmapImage(ByRef clpObject As pdClipboard, ByVal srcIsMeantAsLayer As Boolean) As Boolean
        
    'Unfortunately, a lot of things can go wrong when pasting bitmaps, so we assume failure by default.
    ClipboardPaste_BitmapImage = False
    
    'Verify that the requested data is actually available.  (Hopefully the caller already checked this, but you never know...)
    If clpObject.DoesClipboardHaveBitmapImage() Then
        
        'Next, we want to sort handling by the "priority" bitmap format.  This is the format the caller actually placed on the clipboard
        ' (vs a variant that Windows auto-created to simplify handling).
        Dim priorityFormat As PredefinedClipboardFormatConstants
        priorityFormat = clpObject.GetPriorityBitmapFormat()
                
        'Bitmap formats may require special behavior on the part of external load functions, so it's important that we accurately
        ' mark the module-level tracker with both the current format (what we retrieved from the clipboard; this may have been
        ' auto-generated by Windows), and the original format the caller placed on the clipboard.
        m_ClipboardInfo.pdci_OriginalFormat = priorityFormat
        
        'If DIBv5 is the format the caller actually placed on the clipboard, retrieve it first.  Otherwise, use the CF_DIB data.
        ' (Ignore CF_BITMAP for now, as it would require specialized handling.)
        Dim rawClipboardData() As Byte, successfulExtraction As Boolean
        If priorityFormat = CF_DIBV5 Then
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "DIBv5 selected as priority retrieval format."
            #End If
            successfulExtraction = clpObject.GetClipboardBinaryData(CF_DIBV5, rawClipboardData)
            m_ClipboardInfo.pdci_CurrentFormat = CF_DIBV5
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Generic DIB selected as priority retrieval format."
            #End If
            successfulExtraction = clpObject.GetClipboardBinaryData(CF_DIB, rawClipboardData)
            m_ClipboardInfo.pdci_CurrentFormat = CF_DIB
        End If
        
        'If the extraction was successful, we can use similar handling for both cases
        If successfulExtraction Then
        
            'Perform some failsafe validation on the DIB header
            Dim dibHeaderOkay As Boolean
            dibHeaderOkay = True
            
            'First, make sure we have at least 40 bytes of data to work with.  (Anything smaller than this and we can't even retrieve a header!)
            If UBound(rawClipboardData) < 40 Then dibHeaderOkay = False
            
            'If we have at least 40 bytes of data, copy them into a default BITMAPINFOHEADER.  This struct is shared between regular DIBs
            ' and v5 DIBs.
            Dim bmpHeader As BITMAPINFOHEADER, bmpV5Header As BITMAPV5HEADER
            If dibHeaderOkay Then
                
                'Retrieve a copy of the bitmap's header in standard, 40-byte format.  This gives us some default values like width, height,
                ' and color depth.
                CopyMemory ByVal VarPtr(bmpHeader), ByVal VarPtr(rawClipboardData(0)), LenB(bmpHeader)
                
                'Validate the header size; it must match a default DIB header, or a v5 DIB header
                If (bmpHeader.biSize <> LenB(bmpHeader)) And (bmpHeader.biSize <> LenB(bmpV5Header)) Then dibHeaderOkay = False
                
                'If a v5 header is present, retrieve it as well
                If (priorityFormat = CF_DIBV5) And (bmpHeader.biSize = LenB(bmpV5Header)) And (UBound(rawClipboardData) > LenB(bmpV5Header)) Then
                    CopyMemory ByVal VarPtr(bmpV5Header), ByVal VarPtr(rawClipboardData(0)), LenB(bmpV5Header)
                    
                    'Track some v5 header data at module-level; external functions may request copies of this, to know how to handle alpha.
                    m_ClipboardInfo.pdci_DIBv5AlphaMask = bmpV5Header.biAlphaMask
                    
                End If
                
                'If the header size checks out, validate width/height next
                If dibHeaderOkay Then
                
                    With bmpHeader
                    
                        'Width must be positive
                        If .biWidth < 0 Then dibHeaderOkay = False
                        
                        'For performance reasons, restrict sizes to 2 ^ 16 in either dimension.  This metric is also used by Chrome and Firefox.
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
                        tmpDIB.createBlank bmpHeader.biWidth, Abs(bmpHeader.biHeight), 32, 0, 0
                    Else
                        tmpDIB.createBlank bmpHeader.biWidth, Abs(bmpHeader.biHeight), 24, 0, 0
                    End If
                    
                    'Calculate the offset required to access the pixel data.  (This value is required by the BMP file format, which PD
                    ' uses as a quick intermediary format.)  Note that some offset calculations only apply to the v5 version of the header.
                    Dim pixelOffset As Long
                    
                    With bmpHeader
                        
                        'Always count the header size in the offset
                        pixelOffset = .biSize
                        
                        'If a color table is included, add it to the offset
                        If .biClrUsed > 0 Then
                            pixelOffset = pixelOffset + .biClrUsed * 4
                        Else
                            If .biBitCount <= 8 Then pixelOffset = pixelOffset + 4 * (2 ^ .biBitCount)
                        End If
                        
                        'Bitfields are optional with certain bit-depths; if bitfields are specified, add them too
                        If (.biCompression = 3) Then
                            If (.biBitCount = 16) Then pixelOffset = pixelOffset + 12
                            If (.biBitCount = 32) Then pixelOffset = pixelOffset + 16
                        End If
                        
                    End With
                    
                    'v5 of the BMP header allows for ICC profiles.  These are supposed to be stored AFTER the pixel data, but some software
                    ' is written by idiots (hi!), so perform a failsafe check for out-of-place profiles.
                    If (priorityFormat = CF_DIBV5) And (bmpV5Header.biProfileData <= pixelOffset) Then pixelOffset = pixelOffset + bmpV5Header.biProfileSize
                                        
                    'We now know enough to create a temporary BMP file as a placeholder for the clipboard data.
                    
                    'Place the temporary file in inside the program-specified temp path
                    Dim tmpClipboardFile As String
                    tmpClipboardFile = g_UserPreferences.GetTempPath & "PDClipboard.bmp"
                    
                    'pdFSO is used to ensure Unicode subfolder compatibility
                    Dim cFile As pdFSO
                    Set cFile = New pdFSO
                    If cFile.FileExist(tmpClipboardFile) Then cFile.KillFile tmpClipboardFile
                        
                    'Populate the BMP file header; it's a simple 14-byte, unchanging struct that requires only a magic number,
                    ' a total filesize, and an offset that points at the pixel bits (NOT the BMP file header, or the embedded
                    ' DIB header - the actual pixel bits).
                    Dim bmpFileHeader As BITMAPFILEHEADER
                    With bmpFileHeader
                        .Type = &H4D42
                        .Size = (UBound(rawClipboardData) + 1) + 14
                        .OffBits = pixelOffset + 14
                    End With
                    
                    Dim hFile As Long
                    If cFile.CreateAppendFileHandle(tmpClipboardFile, hFile) Then
                        
                        'To avoid automatic 4-byte struct alignment, we must write out the header manually.
                        cFile.WriteDataToFile hFile, VarPtr(bmpFileHeader.Type), 2&
                        cFile.WriteDataToFile hFile, VarPtr(bmpFileHeader.Size), 4&
            
                        Dim reservedBytes As Long
                        cFile.WriteDataToFile hFile, VarPtr(reservedBytes), 4&
                        cFile.WriteDataToFile hFile, VarPtr(bmpFileHeader.OffBits), 4&
                        
                        'Simply plop the clipboard data into place last, no changes required
                        cFile.WriteDataToFile hFile, VarPtr(rawClipboardData(0)), UBound(rawClipboardData) + 1
                        cFile.CloseFileHandle hFile
                        
                    End If
                        
                    'We can now use PD's standard image load routine to import the temporary file.  Because we don't want the
                    ' load function to use the temporary file name as the image name, we manually supply a filename to suggest
                    ' if the user eventually tries to save the file.
                    Dim sFile() As String, sTitle As String, sFilename As String
                    ReDim sFile(0) As String
                    sFile(0) = tmpClipboardFile
                    sTitle = g_Language.TranslateMessage("Clipboard Image")
                    sFilename = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
                        
                    'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                    If srcIsMeantAsLayer Then
                        Layer_Handler.loadImageAsNewLayer False, sFile(0), sTitle, True
                    Else
                        LoadFileAsNewImage sFile, False, sTitle, sFilename
                    End If
                            
                    'Once the load is complete, be polite and remove the temporary file
                    cFile.KillFile tmpClipboardFile
                        
                    'Check for load failure.  If the most recent pdImages object is inactive, it's a safe assumption that
                    ' the load operation failed.  (This isn't foolproof, especially if the user loads a ton of images,
                    ' and subsequently unloads images in an arbitrary order - but given the rarity of this situation, I'm content
                    ' to temporarily use this technique for predicting failure.)
                    '
                    'TODO: rewrite PD's central image load functions to return pass/fail status.
                    If g_CurrentImage <= UBound(pdImages) Then
                        If Not pdImages(g_CurrentImage) Is Nothing Then
                            If pdImages(g_CurrentImage).IsActive Then
                                Message "Clipboard data imported successfully "
                                ClipboardPaste_BitmapImage = True
                            End If
                        End If
                    End If
                    
                Else
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "WARNING!  ClipboardPaste_BitmapImage failed because the DIB header failed validation.  Paste abandoned."
                    #End If
                End If
                
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  ClipboardPaste_BitmapImage failed because the DIB header is an invalid size.  Paste abandoned."
                #End If
            End If
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  ClipboardPaste_BitmapImage failed to retrieve raw DIB data.  Paste abandoned."
            #End If
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  ClipboardPaste_BitmapImage was called, but the requested data doesn't exist on the clipboard."
        #End If
    End If
    
End Function

'Because OLE drag/drop is so similar to clipboard functionality, I have included that functionality here.
' Data and Effect are passed as-is from the calling function, while intendedTargetIsLayer controls whether the image file(s)
' (if valid) should be loaded as new layers or new images.
Public Function LoadImageFromDragDrop(ByRef Data As DataObject, ByRef Effect As Long, ByVal intendedTargetIsLayer As Boolean) As Boolean

    Dim sFile() As String
    Dim tmpString As String
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'Verify that the object being dragged is some sort of file or file list
    If Data.GetFormat(vbCFFiles) Then
        
        'Copy the filenames into an array
        ReDim sFile(0 To Data.Files.Count) As String
        
        Dim oleFilename
        
        Dim countFiles As Long
        countFiles = 0
                
        For Each oleFilename In Data.Files
            tmpString = oleFilename
            If (Len(tmpString) <> 0) And cFile.FileExist(tmpString) Then
                sFile(countFiles) = tmpString
                countFiles = countFiles + 1
            End If
        Next oleFilename
        
        'Make sure at least one valid, existant file was found
        If countFiles > 0 Then
        
            'Because the OLE drop may include blank strings, verify the size of the array against countFiles
            ReDim Preserve sFile(0 To countFiles - 1) As String
            
            'If an open image exists, pass the list of filenames to LoadImageAsNewLayer, which will load the images one-at-a-time
            Dim i As Long
            
            If (g_OpenImageCount > 0) And intendedTargetIsLayer Then
                For i = 0 To UBound(sFile)
                    Layer_Handler.loadImageAsNewLayer False, sFile(i)
                Next i
            Else
                LoadFileAsNewImage sFile
            End If
            
            LoadImageFromDragDrop = True
            Exit Function
            
        End If
    
    End If
    
    'If the data is not a file list, see if it's a bitmap
    If Data.GetFormat(vbCFBitmap) Then
    
        'Copy the image into an StdPicture object
        Dim tmpPicture As StdPicture
        Set tmpPicture = Data.GetData(vbCFBitmap)
        
        'Create a temporary DIB and copy the temporary StdPicture object into it
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromPicture tmpPicture
        
        'Ask the DIB to write its contents to file in BMP format
        tmpString = g_UserPreferences.GetTempPath & "PD_DragDrop.tmp"
        tmpDIB.writeToBitmapFile tmpString
        
        'Now that the image is saved on the hard drive, we can delete our temporary objects
        Set tmpPicture = Nothing
        tmpDIB.eraseDIB
        Set tmpDIB = Nothing
        
        'Use the standard image load routine to import the temporary file
        Dim sTitle As String, sFilename As String
        ReDim sFile(0) As String
        sFile(0) = tmpString
        sTitle = g_Language.TranslateMessage("Imported Image")
        sFilename = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
        
        'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
        If (g_OpenImageCount > 0) And intendedTargetIsLayer Then
            Layer_Handler.loadImageAsNewLayer False, sFile(0), sTitle
        Else
            LoadFileAsNewImage sFile, False, sTitle, sFilename
        End If
            
        'Be polite and remove the temporary file
        If cFile.FileExist(tmpString) Then cFile.KillFile tmpString
            
        Message "Image imported successfully "
        
        LoadImageFromDragDrop = True
        Exit Function
    
    'If the data is not a file list, see if it's a URL.
    ElseIf Data.GetFormat(vbCFText) Then
    
        Dim tmpDownloadFile As String
        tmpDownloadFile = Trim$(Data.GetData(vbCFText))
        
        If (StrComp(UCase$(Left$(tmpDownloadFile, 7)), "HTTP://", vbBinaryCompare) = 0) Or (StrComp(UCase$(Left$(tmpDownloadFile, 8)), "HTTPS://", vbBinaryCompare) = 0) Or (StrComp(UCase$(Left$(tmpDownloadFile, 6)), "FTP://", vbBinaryCompare) = 0) Then
        
            Message "Image URL found on clipboard.  Attempting to download..."
            
            tmpDownloadFile = FormInternetImport.downloadURLToTempFile(tmpDownloadFile)
            
            'If the download was successful, we can now use the standard image load routine to import the temporary file
            If Len(tmpDownloadFile) <> 0 Then
                
                'Depending on the number of open images, load the clipboard data as a new image or as a new layer in the current image
                If (g_OpenImageCount > 0) And intendedTargetIsLayer Then
                    Layer_Handler.loadImageAsNewLayer False, tmpDownloadFile
                Else
                    ReDim sFile(0) As String
                    sFile(0) = tmpDownloadFile
                    LoadFileAsNewImage sFile, False, , GetFilename(tmpDownloadFile)
                End If
                
                'Delete the temporary file
                If cFile.FileExist(tmpDownloadFile) Then cFile.KillFile tmpDownloadFile
                
                'Exit!
                LoadImageFromDragDrop = True
                Exit Function
            
            Else
            
                'If the download failed, let the user know that hey, at least we tried.
                Message "Image download failed.  Please supply a valid image URL and try again."
                
            End If
            
        End If
    
    End If
    
    'If we made it all the way here, something went horribly wrong
    LoadImageFromDragDrop = False
    Exit Function

End Function
