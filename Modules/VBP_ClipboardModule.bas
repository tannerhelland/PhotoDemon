Attribute VB_Name = "Clipboard_Handler"
'***************************************************************************
'Clipboard Interface
'Copyright ©2001-2014 by Tanner Helland
'Created: 15/April/01
'Last updated: 29/April/14
'Last update: improve reliability of URL parsing from clipboard HTML data
'
'Module for handling all Windows clipboard routines.  Copy and Paste are the real stars; Cut is not included
' (as there is no purpose for it at present), though Empty Clipboard does make an appearance.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'API functions used to extract file names from clipboard data
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function DragQueryFile Lib "shell32" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal iFile As Long, ByVal lpszFile As String, ByVal cch As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long

Private Const CF_HDROP As Long = 15

'Copy the current layer (or composite image, if copyMerged is true) to the clipboard.
' If a selection is active, crop the image to the layer area first.
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
            If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.fixPremultipliedAlpha False
            
        End If
        
    End If
    
    'Copy the temporary DIB to the clipboard, then erase it
    tmpDIB.copyDIBToClipboard
    tmpDIB.eraseDIB
    
End Sub

'Empty the clipboard
Public Sub ClipboardEmpty()
    Clipboard.Clear
End Sub

'Paste an image (e.g. create new image data based on whatever is in the clipboard).
' The parameter "srcIsMeantAsLayer" controls whether the clipboard data is loaded as a new image, or as a new layer in an existing image.
Public Sub ClipboardPaste(ByVal srcIsMeantAsLayer As Boolean)
    
    Dim tmpClipboardFile As String, tmpDownloadFile As String
    Dim sFile(0) As String
    Dim sTitle As String, sFilename As String
        
    'PNGs on the clipboard get preferential treatment, as they preserve transparency data - so check for them first
    Dim clpObject As cCustomClipboard
    Set clpObject = New cCustomClipboard
    
    'See if clipboard data is available in PNG format.  If it is, attempt to load it.
    ' (If successful, this IF block will manually exit the sub upon completion.)
    If clpObject.IsDataAvailableForFormatName(FormMain.hWnd, "PNG") Then
            
        Dim PNGID As Long
        PNGID = clpObject.FormatIDForName(FormMain.hWnd, "PNG")
        
        Dim PNGData() As Byte
        If clpObject.ClipboardOpen(FormMain.hWnd) Then
        
            If clpObject.GetBinaryData(PNGID, PNGData) Then
                
                'Dump the PNG data out to file
                tmpClipboardFile = g_UserPreferences.GetTempPath & "PDClipboard.png"
                
                Dim fileID As Integer
                fileID = FreeFile()
                Open tmpClipboardFile For Binary As #fileID
                    Put #fileID, 1, PNGData
                Close #fileID
                
                'We can now use the standard image load routine to import the temporary file
                sFile(0) = tmpClipboardFile
                sTitle = g_Language.TranslateMessage("Clipboard Image")
                sFilename = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
                
                'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                If srcIsMeantAsLayer Then
                    Layer_Handler.loadImageAsNewLayer False, sFile(0), sTitle
                Else
                    LoadFileAsNewImage sFile, False, sTitle, sFilename
                End If
                    
                'Be polite and remove the temporary file
                If FileExist(tmpClipboardFile) Then Kill tmpClipboardFile
                    
                Message "Clipboard data imported successfully "
                
                clpObject.ClipboardClose
                Exit Sub
        
            End If
        
            clpObject.ClipboardClose
        
        End If
        
    End If
    
    'If no PNG data was found, look for an HTML fragment.  Chrome and Firefox will include an HTML fragment link to any
    ' copied image from within the browser, which we can use to download the image in question.
    ' (If successful, this IF block will manually exit the sub upon completion.)
    If clpObject.IsDataAvailableForFormatName(FormMain.hWnd, "HTML Format") Then
    
        Dim HtmlID As Long
        HtmlID = clpObject.FormatIDForName(FormMain.hWnd, "HTML Format")
        
        If clpObject.ClipboardOpen(FormMain.hWnd) Then
        
            Dim htmlString As String
            If clpObject.GetTextData(HtmlID, htmlString) Then
                
                'Look for an image reference within the HTML snippet
                If InStr(1, htmlString, "<img ", vbTextCompare) > 0 Then
                
                    'Retrieve the full image path, which will be between the first set of quotation marks following the
                    ' "<img src=" statement in the HTML snippet.
                    Dim vbQuoteMark As String
                    vbQuoteMark = """"
                    
                    'Parse out the URL between the img src quotes
                    Dim urlStart As Long, urlEnd As Long
                    urlStart = InStr(InStr(1, htmlString, "<img "), htmlString, "src=", vbTextCompare)
                    urlStart = InStr(urlStart, htmlString, vbQuoteMark, vbBinaryCompare) + 1
                    
                    'The magic number 6 below is calculated as the length of (src="), + 1 to advance to the
                    ' character immediately following the quotation mark.
                    urlEnd = InStr(urlStart + 6, htmlString, vbQuoteMark, vbBinaryCompare)
                    
                    'As a failsafe, make sure a valid URL was actually found
                    If (urlStart > 0) And (urlEnd > 0) Then
                    
                        Message "Image URL found on clipboard.  Attempting to download..."
                        
                        tmpDownloadFile = FormInternetImport.downloadURLToTempFile(Mid$(htmlString, urlStart, urlEnd - urlStart))
                        
                        'If the download was successful, we can now use the standard image load routine to import the temporary file
                        If Len(tmpDownloadFile) > 0 Then
                        
                            sFile(0) = tmpDownloadFile
                            
                            'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                            If srcIsMeantAsLayer Then
                                Layer_Handler.loadImageAsNewLayer False, sFile(0)
                            Else
                                LoadFileAsNewImage sFile, False
                            End If
                            
                            'Delete the temporary file
                            If FileExist(tmpDownloadFile) Then Kill tmpDownloadFile
                            
                            Message "Clipboard data imported successfully "
                    
                            clpObject.ClipboardClose
                            Exit Sub
                        
                        Else
                        
                            'If the download failed, let the user know that hey, at least we tried.
                            Message "Image download failed.  Please copy a valid image URL to the clipboard and try again."
                            
                        End If
                    
                    End If
                
                End If
                
        
            End If
        
            clpObject.ClipboardClose
        
        End If
    
    End If
    
    
    'Make sure the clipboard format is a bitmap
    If Clipboard.GetFormat(vbCFBitmap) Then
        
        'Copy the image into an StdPicture object
        Dim tmpPicture As StdPicture
        Set tmpPicture = Clipboard.GetData(2)
        
        'Create a temporary DIB and copy the temporary StdPicture object into it
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromPicture tmpPicture
        
        'Ask the DIB to write its contents to file in BMP format
        tmpClipboardFile = g_UserPreferences.GetTempPath & "PD_Clipboard.tmp"
        tmpDIB.writeToBitmapFile tmpClipboardFile
        
        'Now that the image is saved on the hard drive, we can delete our temporary objects
        Set tmpPicture = Nothing
        tmpDIB.eraseDIB
        Set tmpDIB = Nothing
        
        'Use the standard image load routine to import the temporary file
        sFile(0) = tmpClipboardFile
        sTitle = g_Language.TranslateMessage("Clipboard Image")
        sFilename = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
        
        'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
        If srcIsMeantAsLayer Then
            Layer_Handler.loadImageAsNewLayer False, sFile(0), sTitle
        Else
            LoadFileAsNewImage sFile, False, sTitle, sFilename
        End If
            
        'Be polite and remove the temporary file
        If FileExist(tmpClipboardFile) Then Kill tmpClipboardFile
            
        Message "Clipboard data imported successfully "
    
    'Next, see if the clipboard contains text.  If it does, it may be a hyperlink - if so, try and load it.
    ' TODO: make hyperlinks work with "Paste as new layer".  Right now they will always default to loading as a new image.
    ElseIf Clipboard.GetFormat(vbCFText) Then
        
        tmpDownloadFile = Trim$(Clipboard.GetText)
        
        If (StrComp(Left$(tmpDownloadFile, 7), "http://", vbTextCompare) = 0) Or (StrComp(Left$(tmpDownloadFile, 6), "ftp://", vbTextCompare) = 0) Then
        
            Message "Image URL found on clipboard.  Attempting to download..."
            
            tmpDownloadFile = FormInternetImport.downloadURLToTempFile(tmpDownloadFile)
            
            'If the download was successful, we can now use the standard image load routine to import the temporary file
            If Len(tmpDownloadFile) > 0 Then
            
                sFile(0) = tmpDownloadFile
                
                'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
                If srcIsMeantAsLayer Then
                    Layer_Handler.loadImageAsNewLayer False, sFile(0)
                Else
                    LoadFileAsNewImage sFile, False, , getFilename(tmpDownloadFile)
                End If
                
                'Delete the temporary file
                If FileExist(tmpDownloadFile) Then Kill tmpDownloadFile
                
                Message "Clipboard data imported successfully "
        
                clpObject.ClipboardClose
                Exit Sub
            
            Else
            
                'If the download failed, let the user know that hey, at least we tried.
                Message "Image download failed.  Please copy a valid image URL to the clipboard and try again."
                
            End If
            
        End If
        
    'Next, see if the clipboard contains one or more files.  If it does, try to load them.
    ElseIf Clipboard.GetFormat(vbCFFiles) Then
    
        Dim listFiles() As String
        listFiles = ClipboardGetFiles()
        
        'Depending on the request, load the clipboard data as a new image or as a new layer in the current image
        If srcIsMeantAsLayer Then
            Dim i As Long
            
            For i = 0 To UBound(listFiles)
                Layer_Handler.loadImageAsNewLayer False, listFiles(i)
            Next i
            
        Else
            LoadFileAsNewImage listFiles
        End If
        
    Else
        pdMsgBox "The clipboard is empty or it does not contain a valid picture format.  Please copy a valid image onto the clipboard and try again.", vbExclamation + vbOKOnly + vbApplicationModal, "Windows Clipboard Error"
    End If
    
End Sub

'The code in the function below is a heavily modified version of code originally located at:
' http://www.vb-helper.com/howto_track_clipboard.html (link still good as of 21 December '12)
' Many thanks to the original author(s).
Public Function ClipboardGetFiles() As String()
    
    Dim drop_handle As Long
    Dim num_file_names As Long
    Dim file_names() As String
    Dim file_name As String * 1024
    Dim i As Long

    ' Make sure there is file data.
    If Clipboard.GetFormat(vbCFFiles) Then
        
        ' File data exists. Get it.
        ' Open the clipboard.
        If OpenClipboard(0) Then
            ' The clipboard is open.

            ' Get the handle to the dropped list of files.
            drop_handle = GetClipboardData(CF_HDROP)

            ' Get the number of dropped files.
            num_file_names = DragQueryFile(drop_handle, -1, vbNullString, 0)

            ' Get the file names.
            ReDim file_names(0 To num_file_names - 1) As String
            For i = 0 To num_file_names - 1
                ' Get the file name.
                DragQueryFile drop_handle, i, file_name, Len(file_name)

                ' Truncate at the NULL character.
                file_names(i) = Left$(file_name, InStr(file_name, vbNullChar) - 1)
            Next i

            ' Close the clipboard.
            CloseClipboard

            ' Assign the return value.
            ClipboardGetFiles = file_names
            
        End If
        
    End If
    
End Function

'Because OLE drag/drop is so similar to clipboard functionality, I have included that functionality here.
' Data and Effect are passed as-is from the calling function, while intendedTargetIsLayer controls whether the image file(s)
' (if valid) should be loaded as new layers or new images.
Public Function loadImageFromDragDrop(ByRef Data As DataObject, ByRef Effect As Long, ByVal intendedTargetIsLayer As Boolean) As Boolean

    Dim sFile() As String
    Dim tmpString As String
    
    'Verify that the object being dragged is some sort of file or file list
    If Data.GetFormat(vbCFFiles) Then
        
        'Copy the filenames into an array
        ReDim sFile(0 To Data.Files.Count) As String
        
        Dim oleFilename
        
        Dim countFiles As Long
        countFiles = 0
        
        For Each oleFilename In Data.Files
            tmpString = oleFilename
            If (Len(tmpString) > 0) And FileExist(tmpString) Then
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
            
            loadImageFromDragDrop = True
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
        If FileExist(tmpString) Then Kill tmpString
            
        Message "Image imported successfully "
        
        loadImageFromDragDrop = True
        Exit Function
    
    'If the data is not a file list, see if it's a URL.
    ElseIf Data.GetFormat(vbCFText) Then
    
        Dim tmpDownloadFile As String
        tmpDownloadFile = Trim$(Data.GetData(vbCFText))
        
        If (StrComp(Left$(tmpDownloadFile, 7), "http://", vbTextCompare) = 0) Or (StrComp(Left$(tmpDownloadFile, 6), "ftp://", vbTextCompare) = 0) Then
        
            Message "Image URL found on clipboard.  Attempting to download..."
            
            tmpDownloadFile = FormInternetImport.downloadURLToTempFile(tmpDownloadFile)
            
            'If the download was successful, we can now use the standard image load routine to import the temporary file
            If Len(tmpDownloadFile) > 0 Then
                
                'Depending on the number of open images, load the clipboard data as a new image or as a new layer in the current image
                If (g_OpenImageCount > 0) And intendedTargetIsLayer Then
                    Layer_Handler.loadImageAsNewLayer False, tmpDownloadFile
                Else
                    ReDim sFile(0) As String
                    sFile(0) = tmpDownloadFile
                    LoadFileAsNewImage sFile, False, , getFilename(tmpDownloadFile)
                End If
                
                'Delete the temporary file
                If FileExist(tmpDownloadFile) Then Kill tmpDownloadFile
                
                'Exit!
                loadImageFromDragDrop = True
                Exit Function
            
            Else
            
                'If the download failed, let the user know that hey, at least we tried.
                Message "Image download failed.  Please supply a valid image URL and try again."
                
            End If
            
        End If
    
    End If
    
    'If we made it all the way here, something went horribly wrong
    loadImageFromDragDrop = False
    Exit Function

End Function
