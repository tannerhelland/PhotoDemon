Attribute VB_Name = "pngnq_Interface"
'***************************************************************************
'pngnq-s9 Interface
'Copyright ©2011-2013 by Tanner Helland
'Created: 19/December/12
'Last updated: 26/December/12
'Last update: added mechanism for version checking
'
'Module for handling all pngnq-s9 interfacing.  This module is pointless without the accompanying
' pngnq-s9 plugin, which will be in the Data/Plugins subdirectory as "pngnq-s9.exe"
'
'pngnq-s9 is a modified, much-improved variant of the free, open-source pngnq tool.  You can learn more
' about the original pngnq at:
'
' http://pngnq.sourceforge.net/
'
'...and more about pngnq-s9 specifically at:
'
' http://sourceforge.net/projects/pngnqs9/
'
'This project was designed against v2.0.1 of the pngnq-s9 tool (16 Oct '12).  It may not work with
' other versions of the tool.  Additional documentation regarding the use of pngnq-s9 is available
' as part of the official pngnq-s9 zip, downloadable from http://sourceforge.net/projects/pngnqs9/files/
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Is pngnq-s9 available as a plugin?  (NOTE: this is now determined separately from pngnqEnabled.)
Public Function isPngnqAvailable() As Boolean
    If FileExist(g_PluginPath & "pngnq-s9.exe") Then isPngnqAvailable = True Else isPngnqAvailable = False
End Function

'Retrieve the pngnq-s9 version from the file.  There is currently not a good way to do this, so just report the
' expected version (as it's incredibly unlikely that the user will attempt to use another version).
Public Function getPngnqVersion() As String

    If Not isPngnqAvailable Then
        getPngnqVersion = ""
        Exit Function
    Else
        getPngnqVersion = "2.0.1.0"
    End If
    
End Function
