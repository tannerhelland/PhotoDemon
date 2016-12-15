VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Resource Manager
'Copyright 2016-2016 by Tanner Helland
'Created: 13/December/16
'Last updated: 13/December/16
'Last update: initial build
'
'PhotoDemon needs to include a whole swatch of custom resources.  These resources take up a lot of space,
' and we also need to pull different resources depending on things like screen DPI.  To simplify this process,
' we manage resources manually, rather than relying on an external .rc file to do it for us.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Called early in PD's startup, this function checks to see if we've already extracted PD's resource file to a
' separate working folder.  If we have, we grab the data from there.  (This is helpful during debugging.)
' In production builds, the initial collection is pulled from the .exe itself.
Public Function LoadInitialResourceCollection() As Boolean

    'TODO!  Someday we'll build the (finished) resource file directly into the .exe, but for now,
    ' load it from a static debug folder.

End Function
