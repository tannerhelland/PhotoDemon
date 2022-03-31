VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdXCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon XCF (GIMP native file format) Container and Parser
'Copyright 2022-2022 by Tanner Helland
'Created: 30/March/22
'Last updated: 30/March/22
'Last update: initial build
'
'This class (and its associated pdXCF- child classes) handle GIMP XCF parsing duties.  It is custom-built
' for PhotoDemon, with an emphasis on performance and proper color-management of all imported data.
'
'As of 2022 this class is under heavy construction, so please be patient with feature support.
'
'As with all 3rd-party XCF engines, GIMP has many features that don't have direct analogs in PhotoDemon.
' Many of these features are still extracted by this class, but they will not "appear" in the final image.
' My ongoing goal is to expand support in this class as various GIMP features are implemented in PD itself.
'
'Unless otherwise noted, all code in this class is my original work.  I've based my work off the
' "official" XCF spec at this URL (link good as of March 2022):
' https://gitlab.gnome.org/GNOME/gimp/-/blob/master/devel-docs/specifications/xcf.txt
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output;
' this will dump a variety of diagnostic information to the debug log.
Private Const XCF_DEBUG_VERBOSE As Boolean = True

'XCF-specific enums are split between this class and the child pdXCFLayer class
Private Enum xcf_ColorMode
    xcf_RGB = 0
    xcf_Grayscale = 1
    xcf_Indexed = 2
End Enum

#If False Then
    Private Const xcf_RGB = 0, xcf_Grayscale = 1, xcf_Indexed = 2
#End If

Private Enum xcf_Precision
    xcf_08bitIntLinear = 0
    xcf_08bitIntGamma = 1
    xcf_16bitIntLinear = 2
    xcf_16bitIntGamma = 3
    xcf_32bitIntLinear = 4
    xcf_32bitIntGamma = 5
    xcf_16bitFltLinear = 6
    xcf_16bitFltGamma = 7
    xcf_32bitFltLinear = 8
    xcf_32bitFltGamma = 9
    xcf_64bitFltLinear = 10
    xcf_64bitFltGamma = 11
End Enum

#If False Then
    Private Const xcf_08bitIntLinear = 0, xcf_08bitIntGamma = 0, xcf_16bitIntLinear = 0, xcf_16bitIntGamma = 0, xcf_32bitIntLinear = 0, xcf_32bitIntGamma = 0, xcf_16bitFltLinear = 0, xcf_16bitFltGamma = 0, xcf_32bitFltLinear = 0, xcf_32bitFltGamma = 0, xcf_64bitFltLinear = 0, xcf_64bitFltGamma = 0
#End If

'Canvas width/height, color model, and precision all come from the file header
Private m_imageWidth As Long, m_imageHeight As Long
Private m_imageColorMode As xcf_ColorMode, m_imagePrecision As xcf_Precision

'XCF version is critical to correct parsing; different versions have different fields in different places!
Private m_xcfVersion As Long

'Byte-by-byte access is provided, as always, by a pdStream instance
Private m_Stream As pdStream

'Validate a source filename as XCF format.  Validation *does* touch the file -
' we must validate a "magic number" in the header.
Friend Function IsFileXCF(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True, Optional ByVal onSuccessLeaveStreamOpen As Boolean = False) As Boolean
    
    Dim potentiallyXCF As Boolean
    potentiallyXCF = Files.FileExists(srcFilename)
    If potentiallyXCF Then potentiallyXCF = (Files.FileLenW(srcFilename) > 26)
    
    'Check extension up front, if requested.
    If (potentiallyXCF And requireValidFileExtension) Then
        potentiallyXCF = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "xcf", True)
    End If
    
    'Proceed with deeper validation as necessary
    If potentiallyXCF Then
        
        'Attempt to load the file
        Set m_Stream = New pdStream
        If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename) Then
            
            'The first 9 bytes of an XCF file must be the ASCII values "gimp xcf "
            potentiallyXCF = (m_Stream.ReadString_ASCII(9) = "gimp xcf ")
            If (potentiallyXCF And XCF_DEBUG_VERBOSE) Then PDDebug.LogAction "Valid XCF file found"
            
        End If
        
    End If
    
    IsFileXCF = potentiallyXCF
    If (Not IsFileXCF) Or (Not onSuccessLeaveStreamOpen) Then Set m_Stream = Nothing
    
End Function

'Validate and load a candidate QOI file
Friend Function LoadXCF_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME As String = "LoadXCF_FromFile"
    LoadXCF_FromFile = False
    
    'Validate the file
    If Me.IsFileXCF(srcFile, False, True) Then
        
        'Validation only checks the first 9-bytes of the file for a magic ASCII string.
        ' If we're still here, that validation string passed. (We can still reject the file
        ' if the header contains invalid members.)
        
        'Magic number is followed by version (as a string).  We remap this to a long for convenience.
        Dim strVersion As String
        strVersion = LCase$(m_Stream.ReadString_ASCII(4))
        
        If (strVersion = "file") Then
            m_xcfVersion = 0
        Else
            m_xcfVersion = CLng(Right$(strVersion, 3))
        End If
        
        If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "XCF version is " & m_xcfVersion
        
        '1-byte null-padding follows
        m_Stream.ReadByte
        
        'Canvas width/height
        m_imageWidth = m_Stream.ReadLong_BE
        m_imageHeight = m_Stream.ReadLong_BE
        
        'Color model
        m_imageColorMode = m_Stream.ReadLong_BE
        
        'If file is v4 or later, a 4-byte precision value follows next.  (For earlier versions,
        ' this field is omitted and a precision of "8-bit gamma integer" is assumed.)
        If (m_xcfVersion >= 4) Then
            
            Dim tmpInt As Long
            tmpInt = m_Stream.ReadLong_BE
            
            'This field is mapped weirdly; see the spec for details
            Select Case tmpInt
                
                Case 100
                    m_imagePrecision = xcf_08bitIntLinear
                Case 150
                    m_imagePrecision = xcf_08bitIntGamma
                Case 200
                    m_imagePrecision = xcf_16bitIntLinear
                Case 250
                    m_imagePrecision = xcf_16bitIntGamma
                Case 300
                    m_imagePrecision = xcf_32bitIntLinear
                Case 350
                    m_imagePrecision = xcf_32bitIntGamma
                Case 500
                    m_imagePrecision = xcf_16bitFltLinear
                Case 550
                    m_imagePrecision = xcf_16bitFltGamma
                Case 600
                    m_imagePrecision = xcf_32bitFltLinear
                Case 650
                    m_imagePrecision = xcf_32bitFltGamma
                Case 700
                    m_imagePrecision = xcf_64bitFltLinear
                Case 750
                    m_imagePrecision = xcf_64bitFltGamma
                Case Else
                    InternalError FUNC_NAME, "unknown image precision: " & tmpInt
                    Exit Function
            End Select
            
        Else
            m_imagePrecision = xcf_08bitIntGamma
        End If
        
        'And that's it for the header!  Do a few basic validations before continuing.
        If (m_imageWidth <= 0) Or (m_imageHeight <= 0) Then
            InternalError FUNC_NAME, "bad image width/height"
            Exit Function
            
        ElseIf (m_imageWidth >= 65535) Or (m_imageHeight >= 65535) Then
            InternalError FUNC_NAME, "image width/height too big"
            Exit Function
            
        ElseIf (m_imageColorMode < 0) Or (m_imageColorMode > 2) Then
            InternalError FUNC_NAME, "bad color mode: " & m_imageColorMode
            Exit Function
        End If
        
        If XCF_DEBUG_VERBOSE Then PDDebug.LogAction "XCF is " & m_imageWidth & "x" & m_imageHeight & " with color mode " & m_imageColorMode & " and precision " & m_imagePrecision
        
        'Still here?  Header looks okay.  Time to proceed with image properties!
        
    'No penalty on failed validation; exit immediately
    Else
        Exit Function
    End If

End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdXCF." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdXCF." & funcName & "() reported an error: " & errDescription
    End If
End Sub

