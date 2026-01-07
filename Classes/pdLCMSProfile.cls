VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdLCMSProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon LCMS (LittleCMS) Color Profile Manager
'Copyright 2015-2026 by Tanner Helland
'Created: 26/May/16
'Last updated: 09/June/16
'Last update: add support for linear RGB profiles
'
'This class interacts heavily with the LittleCMS module.  Look there for implementation details.
' (Also, this class obviously relies on the LittleCMS DLL, which must be present inside PhotoDemon's Plugins folder.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Once an LCMS profile has been successfully created, this value will be non-zero
Private m_ProfileHandle As Long

Friend Function CreateFromFile(ByRef srcFile As String) As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
        m_ProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(cStream.Peek_PointerOnly(0, Files.FileLenW(srcFile)), Files.FileLenW(srcFile))
        cStream.StopStream
    End If
    CreateFromFile = (m_ProfileHandle <> 0)
End Function

Friend Function CreateFromPointer(ByVal ptrToICCProfile As Long, ByVal iccProfileSizeInBytes As Long) As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(ptrToICCProfile, iccProfileSizeInBytes)
    CreateFromPointer = (m_ProfileHandle <> 0)
End Function

Friend Function CreateFromPDICCObject(ByRef srcProfile As pdICCProfile) As Boolean
    
    CreateFromPDICCObject = True
    
    'Apply a few failsafe checks before attempting to create a profile
    If (srcProfile Is Nothing) Then CreateFromPDICCObject = False
    If CreateFromPDICCObject Then CreateFromPDICCObject = srcProfile.HasICCData
    If CreateFromPDICCObject Then
        If (m_ProfileHandle <> 0) Then ReleaseProfile
        CreateFromPDICCObject = Me.CreateFromPointer(srcProfile.GetICCDataPointer, srcProfile.GetICCDataSize)
    End If
    
End Function

Friend Function CreateGenericGrayscaleProfile(Optional ByVal useGamma As Double = 1#) As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadStockGrayProfile(useGamma)
    CreateGenericGrayscaleProfile = (m_ProfileHandle <> 0)
End Function

Friend Function CreateCustomRGBProfile(ByVal ptrToWhitePointxyY As Long, ByVal ptrTo3xyYPrimaries As Long, Optional ByVal gammaCorrectFactor As Double = 1#) As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadCustomRGBProfile(ptrToWhitePointxyY, ptrTo3xyYPrimaries, gammaCorrectFactor)
    CreateCustomRGBProfile = (m_ProfileHandle <> 0)
End Function

Friend Function CreateCustomRGBProfile_Advanced(ByVal ptrToWhitePointxyY As Long, ByVal ptrTo3xyYPrimaries As Long, ByVal pSrcToneCurveR As Long, ByVal pSrcToneCurveG As Long, ByVal pSrcToneCurveB As Long, Optional ByVal freeCurveForMe As Boolean = True) As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadCustomRGBProfile_Advanced(ptrToWhitePointxyY, ptrTo3xyYPrimaries, pSrcToneCurveR, pSrcToneCurveG, pSrcToneCurveB, freeCurveForMe)
    CreateCustomRGBProfile_Advanced = (m_ProfileHandle <> 0)
End Function

Friend Function CreateLinearRGBProfile() As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadLinearRGBProfile()
    CreateLinearRGBProfile = (m_ProfileHandle <> 0)
End Function

Friend Function CreateSRGBProfile(Optional ByVal useIccV4 As Boolean = True) As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadStockSRGBProfile(useIccV4)
    CreateSRGBProfile = (m_ProfileHandle <> 0)
End Function

Friend Function CreateLabProfile(Optional ByVal useVersion4 As Boolean = True) As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadStockLabProfile(useVersion4)
    CreateLabProfile = (m_ProfileHandle <> 0)
End Function

Friend Function GetColorSpace() As LCMS_PROFILE_COLOR_SPACE
    If (m_ProfileHandle <> 0) Then GetColorSpace = LittleCMS.LCMS_GetProfileColorSpace(m_ProfileHandle)
End Function

Friend Function GetConnectionSpace() As LCMS_PROFILE_COLOR_SPACE
    If (m_ProfileHandle <> 0) Then GetConnectionSpace = LittleCMS.LCMS_GetProfileConnectionSpace(m_ProfileHandle)
End Function

Friend Function GetHandle() As Long
    GetHandle = m_ProfileHandle
End Function

Friend Function GetInfoString(ByVal infoType As LCMS_INFOTYPE) As String
    If (m_ProfileHandle <> 0) Then GetInfoString = LittleCMS.LCMS_GetProfileInfoString(m_ProfileHandle, infoType)
End Function

'Retrieve the raw ICC profile bytes; return value is the length of the profile, in bytes.
Friend Function GetRawProfileBytes(ByRef dstBytes() As Byte) As Long
    If (m_ProfileHandle <> 0) Then
        LittleCMS.LCMS_SaveProfileToArray m_ProfileHandle, dstBytes
        GetRawProfileBytes = UBound(dstBytes) + 1
    End If
End Function

Friend Function GetRenderingIntent() As LCMS_RENDERING_INTENT
    If (m_ProfileHandle <> 0) Then GetRenderingIntent = LittleCMS.LCMS_GetProfileRenderingIntent(m_ProfileHandle)
End Function

Friend Function HasProfile() As Boolean
    HasProfile = (m_ProfileHandle <> 0)
End Function

Friend Function IsGrayProfile() As Boolean
    IsGrayProfile = (Me.GetColorSpace = cmsSigGray)
End Function

Friend Function ReleaseProfile() As Boolean
    ReleaseProfile = LittleCMS.LCMS_CloseProfileHandle(m_ProfileHandle)
    If ReleaseProfile Then m_ProfileHandle = 0
End Function

Private Sub Class_Terminate()
    If (m_ProfileHandle <> 0) Then Me.ReleaseProfile
End Sub
