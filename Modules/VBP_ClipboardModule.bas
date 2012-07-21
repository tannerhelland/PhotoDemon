Attribute VB_Name = "ClipboardFunctions"
'***************************************************************************
'Clipboard Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 07/June/12
'Last update: previously this routine only allowed pasting bitmaps, but there's no reason it shouldn't
'             also support WMF/EMF pasting.  So I added support for those.
'
'Module for handling all Windows clipboard routines.  Copy and Paste are the real stars; I include Cut
' for completeness' sake, though it really serves no purpose, and there is also Empty Clipboard
' functionality.
'
'***************************************************************************

Option Explicit

Private Const CLIPBOARD_FORMAT_BMP As Long = 2
Private Const CLIPBOARD_FORMAT_METAFILE As Long = 3

'Copy image
Public Sub ClipboardCopy()
    Clipboard.Clear
    Clipboard.SetData FormMain.ActiveForm.BackBuffer.Image, CLIPBOARD_FORMAT_BMP
End Sub

'Cut image (a stupid command given the current nature of the program, but I include it for completeness' sake)
'Public Sub ClipboardCut()
'    Clipboard.Clear
'    Clipboard.SetData FormMain.ActiveForm.BackBuffer.Image, CLIPBOARD_FORMAT_BMP
'    FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
'    PrepareViewport
'End Sub

'Empty the clipboard
Public Sub ClipboardEmpty()
    Clipboard.Clear
End Sub

'Paste an image (e.g. create a new image based on data in the clipboard
Public Sub ClipboardPaste()
    
    'Make sure the clipboard format is a bitmap or metafile
    If (Clipboard.GetFormat(CLIPBOARD_FORMAT_BMP) = True) Or (Clipboard.GetFormat(CLIPBOARD_FORMAT_METAFILE) = True) Then
        
        FixScrolling = False
        
        'We'll need a temporary form for saving the data to file
        CreateNewImageForm True
        
        FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
        FormMain.ActiveForm.BackBuffer.Picture = Clipboard.GetData(2)
        DoEvents
        
        Dim tmpClipboardFile As String
        tmpClipboardFile = TempPath & "PDClipboard.tmp"
        
        SavePicture FormMain.ActiveForm.BackBuffer.Picture, tmpClipboardFile
        
        Unload FormMain.ActiveForm
        
        'Now that the clipboard data has been saved to a file, we can use the standard load routine to import it
        'Because PreLoadImage requires a string array, create one to send it
        Dim sFile(0) As String
        sFile(0) = tmpClipboardFile
            
        PreLoadImage sFile, False, "Clipboard Image", "Clipboard Image (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
            
        'Be polite and remove the temporary file
        Kill tmpClipboardFile
            
        Message "Clipboard data imported successfully "
    
    Else
        MsgBox "The clipboard is empty or it does not contain a valid picture format.  Please copy a valid image onto the clipboard and try again.", vbCritical + vbOKOnly + vbApplicationModal, "Windows Clipboard Error"
    End If
    
End Sub
