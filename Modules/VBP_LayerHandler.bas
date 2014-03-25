Attribute VB_Name = "Layer_Handler"
'***************************************************************************
'Layer Interface
'Copyright ©2013-2014 by Tanner Helland
'Created: 24/March/14
'Last updated: 24/March/14
'Last update: initial build
'
'This module provides all layer-related functions that interact with PhotoDemon's central processor.  Most of these
' functions are triggered by either the Layer menu, or the Layer toolbox.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'String returned from the common dialog wrapper
    'Dim sFile() As String
    
    'If PhotoDemon_OpenImageDialog(sFile, getModalOwner().hWnd) Then LoadFileAsNewImage sFile

    'Erase sFile
