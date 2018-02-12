Attribute VB_Name = "Plugin_PNGQuant"
'***************************************************************************
'PNGQuant Interface (formerly pngnq-s9 interface)
'Copyright 2012-2018 by Tanner Helland
'Created: 19/December/12
'Last updated: 02/July/14
'Last update: migrate all plugin support to the official pngquant library.  Work on pngnq-s9 has pretty much
'              evaporated since late 2012, so pngquant is the new workhorse for PD's specialized PNG needs.
'
'Module for handling all pngquant interfacing.  This module is pointless without the accompanying
' pngquant plugin, which will be in the App/PhotoDemon/Plugins subdirectory as "pngquant.exe"
'
'pngquant is a free, open-source lossy PNG compression library.  You can learn more about it here:
'
' http://pngquant.org/
'
'PhotoDemon has been designed against v2.5.2.  You should be able to replace PD's pngquant.exe copy with any
' valid 2.x release, including your own custom-compiled copy, without trouble.  Additional documentation regarding
' the use of pngquant is available as part of the official pngquant library, downloadable from http://pngquant.org/.
'
'pngquant was originally available under a BSD license.  I believe subsequent releases have changed to a
' dual-license with GPLv3.  Please see the App/PhotoDemon/Plugins/pngquant-README.txt file for questions regarding
' copyright or licensing, or for the most up-to-date information, go straight to http://pngquant.org/.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Retrieve the PNGQuant plugin version.  Shelling the executable with the "--version" tag will cause it to return
' the current version (and compile date) over stdout.
Public Function GetPngQuantVersion() As String
    
    GetPngQuantVersion = vbNullString
    
    If PluginManager.IsPluginCurrentlyInstalled(CCP_PNGQuant) Then
        
        Dim pngqPath As String
        pngqPath = PluginManager.GetPluginPath & "pngquant.exe"
        
        Dim outputString As String
        If ShellExecuteCapture(pngqPath, "pngquant.exe --version", outputString) Then
        
            'The output string will be a simple version number and release date, e.g. "2.1.1 (February 2014)".
            ' Split the output by spaces, then retrieve the first entry.
            outputString = Trim$(outputString)
            
            Dim versionParts() As String
            versionParts = Split(outputString, " ")
            GetPngQuantVersion = versionParts(0) & ".0"
            
        End If
        
    End If
    
End Function

'Use pngquant to optimize a PNG file.  By default, a "wait for processing to finish" mechanism is used.
Public Function ApplyPNGQuantToFile_Synchronous(ByVal dstFilename As String, Optional ByVal qualityLevel As Long = 80, Optional ByVal optimizeLevel As Long = 3, Optional ByVal useDithering As Boolean = True, Optional ByVal displayConsoleWindow As Boolean = True) As Boolean
    
    ApplyPNGQuantToFile_Synchronous = False
    
    If g_ImageFormats.pngQuantEnabled Then
        
        'Build a full shell path for the pngquant operation
        Dim shellPath As String
        shellPath = PluginManager.GetPluginPath & "pngquant.exe"
        
        Dim cmdParams As String
        cmdParams = "pngquant.exe "
        'Like JPEGs, quality here is a nebulous measurement.  pngquant wants both a minimum quality (the image will not
        ' be saved if the conversion is worse than this) and a maximum quality (which determines how aggressive it is
        ' about paring down the color tree).  Quality 0 is way better than quality 0 on a JPEG; at Quality ~20 there is
        ' some noticeable dithering, but the image still looks pretty great depending on the color complexity.
        ' We allow pngquant to write a new file regardless of how low it needs to drop quality to compensate.
        cmdParams = cmdParams & "--quality=0-" & CStr(qualityLevel) & " "
        
        'Speed controls the color space and search detail when quantizing the image+alpha data
        cmdParams = cmdParams & "--speed=" & CStr(optimizeLevel) & " "
        
        'Dithering is optional; it generally improves the output significantly, at some cost to file size
        If (Not useDithering) Then cmdParams = cmdParams & "--nofs "
        
        'Force overwrite if a file with that name already exists
        cmdParams = cmdParams & "-f "
        
        'Request the addition of a custom "-8bpp.png" extension; without this, PNGquant will use its own extension
        ' (-fs8.png or -or8.png, depending on the use of dithering)
        cmdParams = cmdParams & "--ext -8bpp.png "
                
        'Verbose output is helpful when debugging
        #If DEBUGMODE = 1 Then
            cmdParams = cmdParams & "-v "
        #End If
        
        'Tell pngquant to stop argument processing here
        cmdParams = cmdParams & "-- "
        
        'Add the filename, then go!
        cmdParams = cmdParams & """" & dstFilename & """"
        
        Message "Using pngquant to optimize the PNG file.  This may take a moment..."
                
        'Before launching the shell, launch a single DoEvents.  This gives us some leeway before Windows marks the program
        ' as unresponsive (relevant since we may have already paused for awhile while writing the PNG file)...
        DoEvents
        
        Dim shellCheck As Boolean
        shellCheck = ShellAndWait(shellPath, cmdParams, displayConsoleWindow)
        
        'If the shell was successful and the image was created successfully, overwrite the original 32bpp save
        ' (from FreeImage) with the newly optimized one (from OptiPNG)
        If shellCheck Then
            
            'If successful, PNGQuant created a new file with the name "filename-8bpp.png".  We need to rename that file
            ' to whatever name the user originally supplied - but only if the 8bpp transformation was successful!
            Dim filenameCheck As String
            filenameCheck = Files.FileGetPath(dstFilename) & Files.FileGetName(dstFilename, True) & "-8bpp.png"
            
            'Make sure both FreeImage and PNGQuant were able to generate valid files, then rewrite the FreeImage one
            ' with the PNGQuant one.
            If Files.FileExists(filenameCheck) And Files.FileExists(dstFilename) Then
                Files.FileReplace dstFilename, filenameCheck
                Message "pngquant optimization successful!"
            Else
            
                'If the original filename's extension was not ".png", pngquant will just cram "-8bpp.png" onto
                ' the existing filename+extension.  Check for this case now.
                If Files.FileExists(dstFilename & "-8bpp.png") And Files.FileExists(dstFilename) Then
                    Files.FileReplace dstFilename, dstFilename & "-8bpp.png"
                    Message "pngquant optimization successful!"
                Else
                    Message "PNGQuant could not write file.  Default 32bpp image was saved instead."
                End If
                
            End If
        
            ApplyPNGQuantToFile_Synchronous = True
            
        End If
        
    End If
    
End Function
