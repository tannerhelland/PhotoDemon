VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdICCProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon ICC (International Color Consortium) Profile Manager
'Copyright 2013-2026 by Tanner Helland
'Created: 05/November/13
'Last updated: 21/November/16
'Last update: allow the class to create profiles directly from arbitrary files
'
'Most image formats support the notion of embedded ICC profiles.  ICC profiles can be used to convert an image
' to its "purest" color representation, taking into account any pecularities of the device that captured the
' image (typically a camera), and the device now being used to display the image (typically a monitor).
'
'ICC profile handling is broken into two parts: extracting the profile from an image, then applying that profile
' to the image.  The extraction step is currently handled via FreeImage or GDI+, while the application step is
' handled by LittleCMS (with a fallback to the built-in Windows ICM if we absolutely have to).  In the future I
' may look at adding ExifTool as a possibly mechanism for extracting the profile, as it provides better support
' for esoteric formats (e.g. GIF, which can technically store ICC profiles, but you never see these in the wild).
'
'Note that this class retains a copy of the ICC profile itself, which is typically a ~1-5k byte array.  When a
' pdDIB class is copied, the attached ICC profile (if any) should also be copied.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'If the attached DIB has been forcibly converted to a neutral working space using this profile, this value will
' be set to TRUE.  Make sure to check this if you need to determine "attached" status vs "converted" status.
Private m_ICCProfileWasApplied As Boolean

'If we are able to successfully load an ICC profile for this image, this will be set to TRUE.
Private m_ICCDataExists As Boolean

'The actual ICC Profile is stored in this array
Private m_ICCData() As Byte

'If this profile was created from a file, the filename will be stored here.  This will be null if the profile
' was created from an arbitrary pointer.
Private m_ICCFilename As String

'Clone another pdICCProfile instance
Friend Function CopyProfile(ByRef srcProfile As pdICCProfile) As Boolean
    CopyProfile = (Not srcProfile Is Nothing)
    If CopyProfile Then CopyProfile = srcProfile.HasICCData()
    If CopyProfile Then CopyProfile = Me.LoadICCFromPtr(srcProfile.GetICCDataSize(), srcProfile.GetICCDataPointer())
End Function

'Retrieve the source render intent from the contained ICC profile
Friend Function GetSourceRenderIntent() As Long

    'Make sure we have been given an ICC profile to work with!
    If (UBound(m_ICCData) = 0) Then
        GetSourceRenderIntent = -1
        Exit Function
    End If
    
    'Render intent is stored in bytes 63-67 of the ICC profile, per the spec.  (http://www.color.org/specification/ICC1v43_2010-12.pdf)
    ' Technically only byte 67 matters, as render intent can only be a value of 0, 1, 2, or 3.
    Dim srcIntent As Long
    If (UBound(m_ICCData) >= 67) Then
        srcIntent = m_ICCData(67)
    Else
        srcIntent = 0
    End If
    
    'Validate the intent; it must fall between 0 and 3.  If it does not, set it to 0 (Perceptual)
    If (srcIntent < 0) Or (srcIntent > 3) Then srcIntent = 0
    
    GetSourceRenderIntent = srcIntent
    
End Function

'If a profile has been successfully applied to the parent image, this will return TRUE.
Friend Function HasProfileBeenApplied() As Boolean
    HasProfileBeenApplied = m_ICCProfileWasApplied
End Function

'If external functions (such as our CMYK conversion code) apply a profile independently,
' they can use this function to mark the image as successfully converted.  This saves us
' the extra work of checking the profile again in PD's primary image import function.
Friend Sub MarkSuccessfulProfileApplication()
    m_ICCProfileWasApplied = True
End Sub

'The external CMYK transform code performs its own special transform using the stored ICC profile; as such, it requires a
' pointer to the stored ICC profile data.
Friend Function GetICCDataPointer() As Long
    GetICCDataPointer = VarPtr(m_ICCData(0))
End Function

Friend Function GetICCDataSize() As Long
    GetICCDataSize = UBound(m_ICCData) + 1
End Function

'If this path was created from a file, the filename is retrievable via this function.  A zero-length string is
' returned if the profile was created from a bare pointer.
Friend Function GetOriginalICCPath() As String
    GetOriginalICCPath = m_ICCFilename
End Function

Friend Function GetProfileBytes(ByRef dstBytes() As Byte, ByRef dstLength As Long) As Boolean
    GetProfileBytes = Me.HasICCData()
    If GetProfileBytes Then
        dstLength = UBound(m_ICCData) + 1
        ReDim dstBytes(0 To UBound(m_ICCData)) As Byte
        CopyMemoryStrict VarPtr(dstBytes(0)), VarPtr(m_ICCData(0)), dstLength
    End If
End Function

Friend Function LoadICCFromPtr(ByVal profileSize As Long, ByVal ptrToProfile As Long) As Boolean
    
    ReDim m_ICCData(0 To profileSize - 1) As Byte
    CopyMemoryStrict VarPtr(m_ICCData(0)), ptrToProfile, profileSize
    
    m_ICCFilename = vbNullString
    m_ICCDataExists = True
    LoadICCFromPtr = True

End Function

Friend Function LoadICCFromFile(ByVal profilePath As String) As Boolean
    
    LoadICCFromFile = False
    
    If Files.FileExists(profilePath) Then LoadICCFromFile = Files.FileLoadAsByteArray(profilePath, m_ICCData)
    m_ICCDataExists = LoadICCFromFile
    If m_ICCDataExists Then m_ICCFilename = profilePath Else m_ICCFilename = vbNullString
    
End Function

Friend Function LoadICCFromLCMSProfile(ByRef srcProfile As pdLCMSProfile) As Boolean
    If (Not srcProfile Is Nothing) Then
        m_ICCFilename = vbNullString
        If srcProfile.HasProfile Then m_ICCDataExists = (srcProfile.GetRawProfileBytes(m_ICCData) <> 0)
        LoadICCFromLCMSProfile = m_ICCDataExists
    End If
End Function

Friend Function IsEqual(ByRef srcProfile As pdICCProfile) As Boolean
    If (srcProfile.GetICCDataSize = Me.GetICCDataSize) Then
        IsEqual = VBHacks.MemCmp(Me.GetICCDataPointer, srcProfile.GetICCDataPointer, Me.GetICCDataSize)
    Else
        IsEqual = False
    End If
End Function

Friend Function HasICCData() As Boolean
    HasICCData = m_ICCDataExists
End Function

Private Sub Class_Initialize()
    m_ICCDataExists = False
    m_ICCProfileWasApplied = False
    ReDim m_ICCData(0) As Byte
End Sub
