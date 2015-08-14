Attribute VB_Name = "Plugin_PNGQuant_Interface"
'***************************************************************************
'PNGQuant Interface (formerly pngnq-s9 interface)
'Copyright 2012-2015 by Tanner Helland
'Created: 19/December/12
'Last updated: 02/July/14
'Last update: migrate all plugin support to the official pngquant library.  Work on pngnq-s9 has pretty much
'              evaporated since late 2012, so pngquant is the new workhorse for PD's specialized PNG needs.
'
'Module for handling all PNGQuant interfacing.  This module is pointless without the accompanying
' PNGQuant plugin, which will be in the App/PhotoDemon/Plugins subdirectory as "pngquant.exe"
'
'PNGQuant is a free, open-source lossy PNG compression library.  You can learn more about it here:
'
' http://pngquant.org/
'
'PhotoDemon has been designed against v2.1.1 (02 July '14).  It may not work with other versions.
' Additional documentation regarding the use of PNGQuant is available as part of the official PNGQuant library,
' downloadable from http://pngquant.org/.
'
'PNGQuant is available under a BSD license.  Please see the App/PhotoDemon/Plugins/pngquant-README.txt file
' for questions regarding copyright or licensing.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Is PNGQuant.exe available on this PC?
Public Function isPngQuantAvailable() As Boolean
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(g_PluginPath & "pngquant.exe") Then isPngQuantAvailable = True Else isPngQuantAvailable = False
    
End Function

'Retrieve the PNGQuant plugin version.  Shelling the executable with the "--version" tag will cause it to return
' the current version (and compile date) over stdout.
Public Function getPngQuantVersion() As String

    If Not isPngQuantAvailable Then
        getPngQuantVersion = ""
        Exit Function
    
    Else
        
        Dim pngqPath As String
        pngqPath = g_PluginPath & "pngquant.exe"
        
        Dim outputString As String
        If ShellExecuteCapture(pngqPath, "pngquant.exe --version", outputString) Then
        
            'The output string will be a simple version number and release date, e.g. "2.1.1 (February 2014)".
            ' Split the output by spaces, then retrieve the first entry.
            outputString = Trim$(outputString)
            
            Dim versionParts() As String
            versionParts = Split(outputString, " ")
            getPngQuantVersion = versionParts(0) & ".0"
            
        Else
            getPngQuantVersion = ""
        End If
        
    End If
    
End Function
