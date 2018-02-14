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
'Copyright 2015-2018 by Tanner Helland
'Created: 26/May/16
'Last updated: 09/June/16
'Last update: add support for linear RGB profiles
'
'This class interacts heavily with the LittleCMS module.  Look there for implementation details.
' (Also, this class obviously relies on the LittleCMS DLL, which must be present inside PhotoDemon's Plugins folder.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Once an LCMS profile has been successfully created, this value will be non-zero
Private m_ProfileHandle As Long

Friend Function CreateFromPointer(ByVal ptrToICCProfile As Long, ByVal iccProfileSizeInBytes As Long) As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(ptrToICCProfile, iccProfileSizeInBytes)
    CreateFromPointer = (m_ProfileHandle <> 0)
End Function

Friend Function CreateFromPDDib(ByRef srcDIB As pdDIB) As Boolean
    
    CreateFromPDDib = False
    
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    
    If (Not (srcDIB Is Nothing)) Then
        If srcDIB.ICCProfile.HasICCData Then
            CreateFromPDDib = Me.CreateFromPointer(srcDIB.ICCProfile.GetICCDataPointer, srcDIB.ICCProfile.GetICCDataSize)
        End If
    End If
    
End Function

Friend Function CreateGenericGrayscaleProfile() As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadStockGrayProfile()
    CreateGenericGrayscaleProfile = (m_ProfileHandle <> 0)
End Function

Friend Function CreateLinearRGBProfile() As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadLinearRGBProfile()
    CreateLinearRGBProfile = (m_ProfileHandle <> 0)
End Function

Friend Function CreateSRGBProfile() As Boolean
    If (m_ProfileHandle <> 0) Then ReleaseProfile
    m_ProfileHandle = LittleCMS.LCMS_LoadStockSRGBProfile()
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

Friend Function GetRenderingIntent() As LCMS_RENDERING_INTENT
    If (m_ProfileHandle <> 0) Then GetRenderingIntent = LittleCMS.LCMS_GetProfileRenderingIntent(m_ProfileHandle)
End Function

Friend Function HasProfile() As Boolean
    HasProfile = (m_ProfileHandle <> 0)
End Function

Friend Function ReleaseProfile() As Boolean
    ReleaseProfile = LittleCMS.LCMS_CloseProfileHandle(m_ProfileHandle)
    If ReleaseProfile Then m_ProfileHandle = 0
End Function

Private Sub Class_Terminate()
    If (m_ProfileHandle <> 0) Then Me.ReleaseProfile
End Sub
