VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdLCMSTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon LCMS (LittleCMS) Transformation Manager
'Copyright 2015-2016 by Tanner Helland
'Created: 26/May/16
'Last updated: 09/June/16
'Last update: add support for RGB modification transforms, which allow adjustment of brightness, contrast, hue, and saturation
'
'This class interacts heavily with the LittleCMS module.  Look there for implementation details.
' (Also, this class obviously relies on the LittleCMS DLL, which must be present inside PhotoDemon's Plugins folder.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Once an LCMS transformation has been successfully created, this value will be non-zero
Private m_TransformHandle As Long

Public Function ApplyTransformToArbitraryMemory(ByVal srcPointer As Long, ByVal dstPointer As Long, ByVal srcScanlineWidth As Long, ByVal dstScanlineWidth As Long, ByVal numScanlines As Long, ByVal scanlineWidthPixels As Long, Optional ByVal reverseScanlines As Boolean = False) As Boolean
    If (m_TransformHandle <> 0) Then
        
        Dim i As Long
        If reverseScanlines Then
            For i = 0 To numScanlines - 1
                LittleCMS.LCMS_TransformArbitraryMemory srcPointer + i * srcScanlineWidth, dstPointer + ((numScanlines - 1) - i) * dstScanlineWidth, scanlineWidthPixels, m_TransformHandle
            Next i
        Else
            For i = 0 To numScanlines - 1
                LittleCMS.LCMS_TransformArbitraryMemory srcPointer + i * srcScanlineWidth, dstPointer + i * dstScanlineWidth, scanlineWidthPixels, m_TransformHandle
            Next i
        End If
        
        ApplyTransformToArbitraryMemory = True
        
    End If
End Function

Friend Function ApplyTransformToPDDib(ByRef targetDIB As pdDIB) As Boolean
    If (Not (targetDIB Is Nothing)) And (m_TransformHandle <> 0) Then
        ApplyTransformToPDDib = LittleCMS.LCMS_ApplyTransformToDIB(targetDIB, m_TransformHandle)
    End If
End Function

Friend Function CreateTwoProfileTransform(ByRef srcProfile As pdLCMSProfile, ByRef dstProfile As pdLCMSProfile, Optional ByVal hInputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal hOutputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = cmsFLAGS_COPY_ALPHA) As Boolean
    If (m_TransformHandle <> 0) Then ReleaseTransform
    m_TransformHandle = LittleCMS.LCMS_CreateTwoProfileTransform(srcProfile.GetHandle, dstProfile.GetHandle, hInputFormat, hOutputFormat, trnsRenderingIntent, trnsFlags)
    CreateTwoProfileTransform = CBool(m_TransformHandle <> 0)
End Function

Friend Function CreateInPlaceTransformForDIB(ByRef srcDIB As pdDIB, ByRef srcProfile As pdLCMSProfile, ByRef dstProfile As pdLCMSProfile, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = cmsFLAGS_COPY_ALPHA) As Boolean
    If (m_TransformHandle <> 0) Then ReleaseTransform
    If (Not (srcDIB Is Nothing)) And (Not (srcProfile Is Nothing)) And (Not (dstProfile Is Nothing)) Then
        m_TransformHandle = LittleCMS.LCMS_CreateInPlaceTransformForDIB(srcProfile.GetHandle, dstProfile.GetHandle, srcDIB, trnsRenderingIntent, trnsFlags)
        CreateInPlaceTransformForDIB = CBool(m_TransformHandle <> 0)
    End If
End Function

'Create an RGB modification transform, for adjusting brightness, contrast, hue, and/or saturation.  A source profile is optional;
' if one isn't passed, sRGB will be assumed.
Friend Function CreateRGBModificationTransform(Optional ByRef srcProfile As pdLCMSProfile = Nothing, Optional ByVal newBrightness As Double = 0#, Optional ByVal newContrast As Double = 1#, Optional ByVal newHue As Double = 0#, Optional ByVal newSaturation As Double = 0#, Optional ByVal srcTemp As Long = 0, Optional ByVal dstTemp As Long = 0, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = cmsFLAGS_COPY_ALPHA) As Boolean
    
    If (m_TransformHandle <> 0) Then ReleaseTransform
    
    'Start by creating an array of 5 individual prfoiles.  I know, it's ridiculous - but here's how the profiles look:
    ' 1) sRGB
    ' 2) Lab
    ' 3) Abstract Lab -> Lab color modification
    ' 4) Abstract Lab -> Lab color modification (yes, the abstract profile needs to be used twice, because it defines both input
    '    and output behavior)
    ' 5) sRGB
    Dim transformProfiles() As Long
    ReDim transformProfiles(0 To 4) As Long
    If (srcProfile Is Nothing) Then transformProfiles(0) = LittleCMS.LCMS_LoadStockSRGBProfile() Else transformProfiles(0) = srcProfile.GetHandle
    transformProfiles(1) = LittleCMS.LCMS_LoadStockLabProfile(True)
    transformProfiles(2) = LittleCMS.LCMS_CreateAbstractBCHSProfile(newBrightness, newContrast, newHue, newSaturation, srcTemp, dstTemp)
    transformProfiles(3) = transformProfiles(2)
    transformProfiles(4) = transformProfiles(0)
    
    'Take that horrible mess of profiles and create a single unified transform from them
    m_TransformHandle = LittleCMS.LCMS_CreateMultiProfileTransform(transformProfiles, 5, , , trnsRenderingIntent, trnsFlags)
    
    'We can now unload our temporary profiles
    If (srcProfile Is Nothing) Then LittleCMS.LCMS_CloseProfileHandle transformProfiles(0)
    LittleCMS.LCMS_CloseProfileHandle transformProfiles(1)
    LittleCMS.LCMS_CloseProfileHandle transformProfiles(2)
    'Profiles (3) and (4) are deliberately skipped, because they are just copies of profile (2) and (0) - see above
    
    CreateRGBModificationTransform = CBool(m_TransformHandle <> 0)
    
End Function

Public Function GetHandle() As Long
    GetHandle = m_TransformHandle
End Function

Public Function HasTransform() As Boolean
    HasTransform = CBool(m_TransformHandle <> 0)
End Function

Public Function ReleaseTransform() As Boolean
    ReleaseTransform = LittleCMS.LCMS_DeleteTransform(m_TransformHandle)
    If ReleaseTransform Then m_TransformHandle = 0
End Function

Private Sub Class_Terminate()
    If (m_TransformHandle <> 0) Then Me.ReleaseTransform
End Sub
