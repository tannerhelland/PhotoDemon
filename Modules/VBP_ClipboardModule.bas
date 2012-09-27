Attribute VB_Name = "Clipboard_Handler"
'***************************************************************************
'Clipboard Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 15/April/01
'Last updated: 06/September/12
'Last update: rewrote copy/paste against the new layer class.
'
'Module for handling all Windows clipboard routines.  Copy and Paste are the real stars; Cut is not included
' (as there is no purpose for it at present), though Empty Clipboard does make an appearance.
'
'***************************************************************************

Option Explicit

Private Const CLIPBOARD_FORMAT_BMP As Long = 2
Private Const CLIPBOARD_FORMAT_METAFILE As Long = 3

'Copy image
Public Sub ClipboardCopy()
    
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 24 Then
        pdImages(CurrentImage).mainLayer.copyLayerToClipboard
    Else
    
        'For transparent images, make a copy, pre-multiply it against white, then copy THAT to the clipboard
        Dim tmpLayer As pdLayer
        Set tmpLayer = New pdLayer
        tmpLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer
        tmpLayer.compositeBackgroundColor 255, 255, 255
        tmpLayer.copyLayerToClipboard
    End If
    
End Sub

'Empty the clipboard
Public Sub ClipboardEmpty()
    Clipboard.Clear
End Sub

'Paste an image (e.g. create a new image based on data in the clipboard
Public Sub ClipboardPaste()
    
    'Make sure the clipboard format is a bitmap
    If Clipboard.GetFormat(CLIPBOARD_FORMAT_BMP) Then
        
        'Copy the image into an StdPicture object
        Dim tmpPicture As StdPicture
        Set tmpPicture = Clipboard.GetData(2)
        
        'Create a temporary layer and copy the temporary StdPicture object into it
        Dim tmpLayer As pdLayer
        Set tmpLayer = New pdLayer
        tmpLayer.CreateFromPicture tmpPicture
        
        'Ask the layer to write its contents to file in BMP format
        Dim tmpClipboardFile As String
        tmpClipboardFile = TempPath & "PDClipboard.tmp"
        tmpLayer.writeToBitmapFile tmpClipboardFile
        
        'Now that the image is saved on the hard drive, we can delete our temporary objects
        Set tmpPicture = Nothing
        tmpLayer.eraseLayer
        Set tmpLayer = Nothing
        
        'Use the standard image load routine to import the temporary file
        Dim sFile(0) As String
        sFile(0) = tmpClipboardFile
            
        PreLoadImage sFile, False, "Clipboard Image", "Clipboard Image (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
            
        'Be polite and remove the temporary file
        If FileExist(tmpClipboardFile) Then Kill tmpClipboardFile
            
        Message "Clipboard data imported successfully "
    
    Else
        MsgBox "The clipboard is empty or it does not contain a valid picture format.  Please copy a valid image onto the clipboard and try again.", vbCritical + vbOKOnly + vbApplicationModal, "Windows Clipboard Error"
    End If
    
End Sub
