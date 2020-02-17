Attribute VB_Name = "Plugin_OptiPNG"
'***************************************************************************
'OptiPNG Interface
'Copyright 2016-2020 by Tanner Helland
'Created: 20/April/16
'Last updated: 20/April/16
'Last update: initial build
'
'Module for handling all OptiPNG interfacing.  This module is pointless without the accompanying
' OptiPNG plugin, which will be in the App/PhotoDemon/Plugins subdirectory as "optipng.exe"
'
'OptiPNG is a free, open-source lossless PNG compression library.  You can learn more about it here:
'
' http://optipng.sourceforge.net/
'
'PhotoDemon has been designed against v0.7.6 (03 April '16).  It may not work with other versions.
' Additional documentation regarding the use of OptiPNG is available as part of the official OptiPNG library,
' downloadable from http://optipng.sourceforge.net/.
'
'OptiPNG is available under the zLib license.  Please see the App/PhotoDemon/Plugins/optipng-LICENSE.txt file
' for questions regarding copyright or licensing.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Retrieve the OptiPNG plugin version.  Shelling the executable with the "-version" tag will cause it to return
' the current version (and compile date) over stdout.
Public Function GetOptiPNGVersion() As String
    
    GetOptiPNGVersion = vbNullString
    
    If PluginManager.IsPluginCurrentlyInstalled(CCP_OptiPNG) Then
        
        Dim pluginPath As String
        pluginPath = PluginManager.GetPluginPath & "optipng.exe"
        
        Dim outputString As String
        If ShellExecuteCapture(pluginPath, "optipng.exe -version", outputString) Then
        
            'The output string is quite large, but the first line will always look like "OptiPNG version 0.7.6".
            ' Split the output by lines, then by spaces, and retrieve the last word of the first line.
            outputString = Trim$(outputString)
            Dim versionLines() As String
            versionLines = Split(outputString, vbCrLf, , vbBinaryCompare)
            
            If VBHacks.IsArrayInitialized(versionLines) Then
                
                Dim versionParts() As String
                versionParts = Split(versionLines(0), " ")
                
                If VBHacks.IsArrayInitialized(versionParts) Then
                    If (UBound(versionParts) >= 2) Then GetOptiPNGVersion = versionParts(2) & ".0"
                End If
            
            End If
            
        End If
        
    End If
    
End Function

'Use OptiPNG to optimize a PNG file.  By default, a "wait for processing to finish" mechanism is used.
Public Function ApplyOptiPNGToFile_Synchronous(ByVal dstFilename As String, Optional ByVal optimizeLevel As Long = 1) As Boolean
    
    ApplyOptiPNGToFile_Synchronous = False
    
    If PluginManager.IsPluginCurrentlyEnabled(CCP_OptiPNG) Then
        
        'Build a full shell path for the pngquant operation
        Dim shellPath As String
        shellPath = PluginManager.GetPluginPath & "optipng.exe"
        
        Dim optimizeFlags As String
        optimizeFlags = "optipng.exe "
        Select Case optimizeLevel
            
            Case 1
                optimizeFlags = "-o1 -nz"
            
            Case 2
                optimizeFlags = "-o1"
            
            Case 3
                optimizeFlags = "-o2"
            
        End Select
        
        'Strip any metadata.  (If the user requested custom metadata embedding, we will apply it in a subsequent step.)
        optimizeFlags = optimizeFlags & " -strip all "
        
        'Add the target filename
        optimizeFlags = optimizeFlags & " """ & dstFilename & """"
        
        Message "Using OptiPNG to optimize the PNG file.  This may take a moment..."
                
        'Before launching the shell, launch a single DoEvents.  This gives us some leeway before Windows marks the program
        ' as unresponsive (relevant since we may have already paused for awhile while writing the PNG file)...
        DoEvents
        
        Dim shellCheck As Boolean
        shellCheck = Files.ShellAndWait(shellPath, optimizeFlags, False)
        
        'If the shell was successful and the image was created successfully, overwrite the original 32bpp save
        ' (from FreeImage) with the newly optimized one (from OptiPNG)
        If shellCheck Then
            Message "OptiPNG optimization successful!"
            ApplyOptiPNGToFile_Synchronous = True
        End If
        
    End If
    
End Function
