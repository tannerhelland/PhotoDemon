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
'Last updated: 26/May/16
'Last update: migrate existing LCMS code to this more OOP-friendly design
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

Public Function ApplyTransformToArbitraryMemory(ByVal srcPointer As Long, ByVal dstPointer As Long, ByVal srcScanlineWidth As Long, ByVal dstScanlineWidth As Long, ByVal numScanlines As Long, ByVal scanlineWidthPixels As Long) As Boolean
    If (m_TransformHandle <> 0) Then
        
        Dim i As Long
        For i = 0 To numScanlines - 1
            LittleCMS.LCMS_TransformArbitraryMemory srcPointer + i * srcScanlineWidth, dstPointer + i * dstScanlineWidth, scanlineWidthPixels, m_TransformHandle
        Next i
        
        ApplyTransformToArbitraryMemory = True
        
    End If
End Function

Friend Function ApplyTransformToPDDib(ByRef targetDIB As pdDIB) As Boolean
    If (Not (targetDIB Is Nothing)) And (m_TransformHandle <> 0) Then
        ApplyTransformToPDDib = LittleCMS.LCMS_ApplyTransformToDIB(targetDIB, m_TransformHandle)
    End If
End Function

Friend Function CreateTwoProfileTransform(ByRef srcProfile As pdLCMSProfile, ByRef dstProfile As pdLCMSProfile, Optional ByVal hInputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal hOutputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = 0) As Boolean
    If (m_TransformHandle <> 0) Then ReleaseTransform
    m_TransformHandle = LittleCMS.LCMS_CreateTwoProfileTransform(srcProfile.GetHandle, dstProfile.GetHandle, hInputFormat, hOutputFormat, trnsRenderingIntent, trnsFlags)
    CreateTwoProfileTransform = CBool(m_TransformHandle <> 0)
End Function

Friend Function CreateInPlaceTransformForDIB(ByRef srcDIB As pdDIB, ByRef srcProfile As pdLCMSProfile, ByRef dstProfile As pdLCMSProfile, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = 0) As Boolean
    If (m_TransformHandle <> 0) Then ReleaseTransform
    If (Not (srcDIB Is Nothing)) And (Not (srcProfile Is Nothing)) And (Not (dstProfile Is Nothing)) Then
        m_TransformHandle = LittleCMS.LCMS_CreateInPlaceTransformForDIB(srcProfile.GetHandle, dstProfile.GetHandle, srcDIB, trnsRenderingIntent, trnsFlags)
        CreateInPlaceTransformForDIB = CBool(m_TransformHandle <> 0)
    End If
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
