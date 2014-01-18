Attribute VB_Name = "Image_Autosave_Handler"
'***************************************************************************
'Image Autosave Handler
'Copyright ©2013-2014 by Tanner Helland
'Created: 18/January/14
'Last updated: 18/January/14
'Last update: initial build
'
'PhotoDemon's Autosave engine is closely tied to the pdUndo class, so some understanding of that class is necessary
' to appreciate how this module operates.
'
'All Undo/Redo data is saved to the hard drive, in a temp folder of the user's choosing (the Windows temp folder
' by default).  This data is cleared whenever an image is unloaded, and an extra pass is made at program shutdown
' "just to be safe".
'
'In the event of an unclean shutdown, this module searches the temp folder for any PhotoDemon-specific data.  If
' some is found, the user is given a choice to restore those files.  If the user declines, that data is wiped
' (to prevent future unclean shutdown checks from re-detecting it).
'
'As part of its Autosave functionality, this module also handles the creation and subsequent destruction of a
' "clean shutdown" file.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Check to make sure the last program shutdown was clean.  If it was, return TRUE (and write out a new safe shutdown file).
' If it was not, return FALSE.
Public Function wasLastShutdownClean() As Boolean

    Dim safeShutdownPath As String
    safeShutdownPath = g_UserPreferences.getPresetPath & "SafeShutdown.xml"
    
    'If a previous program session terminated unexpectedly, its safe shutdown file will still be present
    If FileExist(safeShutdownPath) Then
    
        wasLastShutdownClean = False

    'The previous shutdown was clean.  Write a new safe shutdown file.
    Else
    
        Dim xmlEngine As pdXML
        Set xmlEngine = New pdXML
        
        xmlEngine.prepareNewXML "Safe shutdown"
        
        xmlEngine.writeBlankLine
        xmlEngine.writeComment "This file is used to see if the previous PhotoDemon session terminated unexpectedly."
        xmlEngine.writeBlankLine
        xmlEngine.writeTag "SessionDate", Format$(Now, "Long Date")
        xmlEngine.writeTag "SessionTime", Format$(Now, "h:mm AMPM")
        xmlEngine.writeBlankLine
        
        xmlEngine.writeXMLToFile safeShutdownPath
        
        wasLastShutdownClean = True
    
    End If
    
    
End Function

'If the program has shut itself down without incident, the last thing it does will be notifying this sub.
' (This sub clears the safe shutdown file.)
Public Sub notifyCleanShutdown()
    
    Dim safeShutdownPath As String
    safeShutdownPath = g_UserPreferences.getPresetPath & "SafeShutdown.xml"
    
    If FileExist(safeShutdownPath) Then Kill safeShutdownPath

End Sub
