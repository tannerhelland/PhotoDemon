VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFITS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon FITS ("Flexible Image Transport System") Decoder
'Copyright 2025-2025 by Tanner Helland
'Created: 19/November/25
'Last updated: 19/November/25
'Last update: initial build
'
'Per Wikipedia, "FITS is the most commonly used digital file format in astronomy":
' https://en.wikipedia.org/wiki/FITS
'
'You can find the FITS spec here (link good as of Nov 2025):
' https://fits.gsfc.nasa.gov/standard40/fits_standard40aa-le.pdf
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Friend Function IsFileFITS(ByRef srcFile As String, Optional ByVal requireValidFileExtension As Boolean = True) As Boolean

    Dim potentialMatch As Boolean
    potentialMatch = Files.FileExists(srcFile)
    
    If potentialMatch And requireValidFileExtension Then
        potentialMatch = Strings.StringsEqualAny(Files.FileGetExtension(srcFile), True, "fits", "fit", "fts")
    End If
    
    If potentialMatch Then
    
    End If
    
End Function
