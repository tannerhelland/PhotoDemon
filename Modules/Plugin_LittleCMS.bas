Attribute VB_Name = "LittleCMS"
'***************************************************************************
'LittleCMS Interface
'Copyright 2016-2016 by Tanner Helland
'Created: 21/April/16
'Last updated: 09/June/16
'Last update: continued feature expansion
'
'Module for handling all LittleCMS interfacing.  This module is pointless without the accompanying
' LittleCMS plugin, which will be in the App/PhotoDemon/Plugins subdirectory as "lcms2.dll".
'
'LittleCMS is a free, open-source color management library.  You can learn more about it here:
'
' http://www.littlecms.com/
'
'PhotoDemon has been designed against v 2.7.0.  It may not work with other versions.
' Additional documentation regarding the use of LittleCMS is available as part of the official LittleCMS library,
' available from https://github.com/mm2/Little-CMS.
'
'LittleCMS is available under the MIT license.  Please see the App/PhotoDemon/Plugins/lcms2-LICENSE.txt file
' for questions regarding copyright or licensing.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Type LCMS_XYZ
    x As Double
    y As Double
    z As Double
End Type

Public Type LCMS_xyY
    x As Double
    y As Double
    YY As Double
End Type

Public Type LCMS_Lab
    l As Double
    a As Double
    b As Double
End Type

Public Type LCMS_LCh
    l As Double
    c As Double
    h As Double
End Type

Public Type LCMS_JCh
    j As Double
    c As Double
    h As Double
End Type


'LCMS allows you to define custom pixel formatters, but they also provide a large collection of pre-formatted values.
' We prefer to use these whenever possible.
Public Enum LCMS_PIXEL_FORMAT
    TYPE_GRAY_8 = &H30009
    TYPE_GRAY_8_REV = &H32009
    TYPE_GRAY_16 = &H3000A
    TYPE_GRAY_16_REV = &H3200A
    TYPE_GRAY_16_SE = &H3080A
    TYPE_GRAYA_8 = &H30089
    TYPE_GRAYA_16 = &H3008A
    TYPE_GRAYA_16_SE = &H3088A
    TYPE_GRAYA_8_PLANAR = &H31089
    TYPE_GRAYA_16_PLANAR = &H3108A
    TYPE_RGB_8 = &H40019
    TYPE_RGB_8_PLANAR = &H41019
    TYPE_BGR_8 = &H40419
    TYPE_BGR_8_PLANAR = &H41419
    TYPE_RGB_16 = &H4001A
    TYPE_RGB_16_PLANAR = &H4101A
    TYPE_RGB_16_SE = &H4081A
    TYPE_BGR_16 = &H4041A
    TYPE_BGR_16_PLANAR = &H4141A
    TYPE_BGR_16_SE = &H40C1A
    TYPE_RGBA_8 = &H40099
    TYPE_RGBA_8_PLANAR = &H41099
    TYPE_ARGB_8_PLANAR = &H45099
    TYPE_ABGR_8_PLANAR = &H41499
    TYPE_BGRA_8_PLANAR = &H45499
    TYPE_RGBA_16 = &H4009A
    TYPE_RGBA_16_PLANAR = &H4109A
    TYPE_RGBA_16_SE = &H4089A
    TYPE_ARGB_8 = &H44099
    TYPE_ARGB_16 = &H4409A
    TYPE_ABGR_8 = &H40499
    TYPE_ABGR_16 = &H4049A
    TYPE_ABGR_16_PLANAR = &H4149A
    TYPE_ABGR_16_SE = &H40C9A
    TYPE_BGRA_8 = &H44499
    TYPE_BGRA_16 = &H4449A
    TYPE_BGRA_16_SE = &H4489A
    TYPE_CMY_8 = &H50019
    TYPE_CMY_8_PLANAR = &H51019
    TYPE_CMY_16 = &H5001A
    TYPE_CMY_16_PLANAR = &H5101A
    TYPE_CMY_16_SE = &H5081A
    TYPE_CMYK_8 = &H60021
    TYPE_CMYKA_8 = &H600A1
    TYPE_CMYK_8_REV = &H62021
    TYPE_YUVK_8 = &H62021
    TYPE_CMYK_8_PLANAR = &H61021
    TYPE_CMYK_16 = &H60022
    TYPE_CMYK_16_REV = &H62022
    TYPE_YUVK_16 = &H62022
    TYPE_CMYK_16_PLANAR = &H61022
    TYPE_CMYK_16_SE = &H60822
    TYPE_KYMC_8 = &H60421
    TYPE_KYMC_16 = &H60422
    TYPE_KYMC_16_SE = &H60C22
    TYPE_KCMY_8 = &H64021
    TYPE_KCMY_8_REV = &H66021
    TYPE_KCMY_16 = &H64022
    TYPE_KCMY_16_REV = &H66022
    TYPE_KCMY_16_SE = &H64822
    TYPE_CMYK5_8 = &H130029
    TYPE_CMYK5_16 = &H13002A
    TYPE_CMYK5_16_SE = &H13082A
    TYPE_KYMC5_8 = &H130429
    TYPE_KYMC5_16 = &H13042A
    TYPE_KYMC5_16_SE = &H130C2A
    TYPE_CMYK6_8 = &H140031
    TYPE_CMYK6_8_PLANAR = &H141031
    TYPE_CMYK6_16 = &H140032
    TYPE_CMYK6_16_PLANAR = &H141032
    TYPE_CMYK6_16_SE = &H140832
    TYPE_CMYK7_8 = &H150039
    TYPE_CMYK7_16 = &H15003A
    TYPE_CMYK7_16_SE = &H15083A
    TYPE_KYMC7_8 = &H150439
    TYPE_KYMC7_16 = &H15043A
    TYPE_KYMC7_16_SE = &H150C3A
    TYPE_CMYK8_8 = &H160041
    TYPE_CMYK8_16 = &H160042
    TYPE_CMYK8_16_SE = &H160842
    TYPE_KYMC8_8 = &H160441
    TYPE_KYMC8_16 = &H160442
    TYPE_KYMC8_16_SE = &H160C42
    TYPE_CMYK9_8 = &H170049
    TYPE_CMYK9_16 = &H17004A
    TYPE_CMYK9_16_SE = &H17084A
    TYPE_KYMC9_8 = &H170449
    TYPE_KYMC9_16 = &H17044A
    TYPE_KYMC9_16_SE = &H170C4A
    TYPE_CMYK10_8 = &H180051
    TYPE_CMYK10_16 = &H180052
    TYPE_CMYK10_16_SE = &H180852
    TYPE_KYMC10_8 = &H180451
    TYPE_KYMC10_16 = &H180452
    TYPE_KYMC10_16_SE = &H180C52
    TYPE_CMYK11_8 = &H190059
    TYPE_CMYK11_16 = &H19005A
    TYPE_CMYK11_16_SE = &H19085A
    TYPE_KYMC11_8 = &H190459
    TYPE_KYMC11_16 = &H19045A
    TYPE_KYMC11_16_SE = &H190C5A
    TYPE_CMYK12_8 = &H1A0061
    TYPE_CMYK12_16 = &H1A0062
    TYPE_CMYK12_16_SE = &H1A0862
    TYPE_KYMC12_8 = &H1A0461
    TYPE_KYMC12_16 = &H1A0462
    TYPE_KYMC12_16_SE = &H1A0C62
    TYPE_XYZ_16 = &H9001A
    TYPE_Lab_8 = &HA0019
    TYPE_ALab_8 = &HA0499
    TYPE_Lab_16 = &HA001A
    TYPE_Yxy_16 = &HE001A
    TYPE_YCbCr_8 = &H70019
    TYPE_YCbCr_8_PLANAR = &H71019
    TYPE_YCbCr_16 = &H7001A
    TYPE_YCbCr_16_PLANAR = &H7101A
    TYPE_YCbCr_16_SE = &H7081A
    TYPE_YUV_8 = &H80019
    TYPE_YUV_8_PLANAR = &H81019
    TYPE_YUV_16 = &H8001A
    TYPE_YUV_16_PLANAR = &H8101A
    TYPE_YUV_16_SE = &H8081A
    TYPE_HLS_8 = &HD0019
    TYPE_HLS_8_PLANAR = &HD1019
    TYPE_HLS_16 = &HD001A
    TYPE_HLS_16_PLANAR = &HD101A
    TYPE_HLS_16_SE = &HD081A
    TYPE_HSV_8 = &HC0019
    TYPE_HSV_8_PLANAR = &HC1019
    TYPE_HSV_16 = &HC001A
    TYPE_HSV_16_PLANAR = &HC101A
    TYPE_HSV_16_SE = &HC081A

    TYPE_NAMED_COLOR_INDEX = &HA&

    TYPE_XYZ_FLT = &H49001C
    TYPE_Lab_FLT = &H4A001C
    TYPE_GRAY_FLT = &H43000C
    TYPE_RGB_FLT = &H44001C
    TYPE_CMYK_FLT = &H460024
    TYPE_XYZA_FLT = &H49009C
    TYPE_LabA_FLT = &H4A009C
    TYPE_RGBA_FLT = &H44009C

    TYPE_XYZ_DBL = &H490018
    TYPE_Lab_DBL = &H4A0018
    TYPE_GRAY_DBL = &H430008
    TYPE_RGB_DBL = &H440018
    TYPE_CMYK_DBL = &H460020
    TYPE_LabV2_8 = &H1E0019
    TYPE_ALabV2_8 = &H1E0499
    TYPE_LabV2_16 = &H1E001A

    TYPE_GRAY_HALF_FLT = &H43000A
    TYPE_RGB_HALF_FLT = &H44001A
    TYPE_RGBA_HALF_FLT = &H44009A
    TYPE_CMYK_HALF_FLT = &H460022

    TYPE_ARGB_HALF_FLT = &H44409A
    TYPE_BGR_HALF_FLT = &H44041A
    TYPE_BGRA_HALF_FLT = &H44449A
    TYPE_ABGR_HALF_FLT = &H44041A
End Enum

'LCMS supports more intents than the default ICC spec does
Public Enum LCMS_RENDERING_INTENT
    INTENT_PERCEPTUAL = 0
    INTENT_RELATIVE_COLORIMETRIC = 1
    INTENT_SATURATION = 2
    INTENT_ABSOLUTE_COLORIMETRIC = 3
    INTENT_PRESERVE_K_ONLY_PERCEPTUAL = 10
    INTENT_PRESERVE_K_ONLY_RELATIVE_COLORIMETRIC = 11
    INTENT_PRESERVE_K_ONLY_SATURATION = 12
    INTENT_PRESERVE_K_PLANE_PERCEPTUAL = 13
    INTENT_PRESERVE_K_PLANE_RELATIVE_COLORIMETRIC = 14
    INTENT_PRESERVE_K_PLANE_SATURATION = 15
End Enum

'When creating transforms, additional flags can be used to modify the transform process
Public Enum LCMS_TRANSFORM_FLAGS
    'Flags
    cmsFLAGS_NOCACHE = &H40&                       ' Inhibit 1-pixel cache
    cmsFLAGS_NOOPTIMIZE = &H100&                   ' Inhibit optimizations
    cmsFLAGS_NULLTRANSFORM = &H200&                ' Don't transform anyway
    ' Proofing flags
    cmsFLAGS_GAMUTCHECK = &H1000&                  ' Out of Gamut alarm
    cmsFLAGS_SOFTPROOFING = &H4000&                ' Do softproofing
    ' Misc
    cmsFLAGS_BLACKPOINTCOMPENSATION = &H2000&
    cmsFLAGS_NOWHITEONWHITEFIXUP = &H4&            ' Don't fix scum dot
    cmsFLAGS_HIGHRESPRECALC = &H400&               ' Use more memory to give better accurancy
    cmsFLAGS_LOWRESPRECALC = &H800&                ' Use less memory to minimize resouces
    ' For devicelink creation
    cmsFLAGS_8BITS_DEVICELINK = &H8&               ' Create 8 bits devicelinks
    cmsFLAGS_GUESSDEVICECLASS = &H20&              ' Guess device class (for transform2devicelink)
    cmsFLAGS_KEEP_SEQUENCE = &H80&                 ' Keep profile sequence for devicelink creation
    ' Specific to a particular optimizations
    cmsFLAGS_FORCE_CLUT = &H2&                     ' Force CLUT optimization
    cmsFLAGS_CLUT_POST_LINEARIZATION = &H1&        ' create postlinearization tables if possible
    cmsFLAGS_CLUT_PRE_LINEARIZATION = &H10&        ' create prelinearization tables if possible
    ' CRD special
    cmsFLAGS_NODEFAULTRESOURCEDEF = &H1000000
    cmsFLAGS_COPY_ALPHA = &H4000000                ' alpha channels are copied on cmsDoTransform()
End Enum

'Only the first eight values (through F8) are actual LCMS defines; the others are provided for reference and interop, only
Public Enum LCMS_ILLUMINANT
    cmsILLUMINANT_TYPE_UNKNOWN = &H0
    cmsILLUMINANT_TYPE_D50 = &H1
    cmsILLUMINANT_TYPE_D65 = &H2
    cmsILLUMINANT_TYPE_D93 = &H3
    cmsILLUMINANT_TYPE_F2 = &H4
    cmsILLUMINANT_TYPE_D55 = &H5
    cmsILLUMINANT_TYPE_A = &H6
    cmsILLUMINANT_TYPE_E = &H7
    cmsILLUMINANT_TYPE_F8 = &H8
End Enum

'Return the current library version as a Long, e.g. "2.7" is returned as "2070"
Private Declare Function cmsGetEncodedCMMversion Lib "lcms2.dll" () As Long

'Error logger registration
Private Declare Sub cmsSetLogErrorHandler Lib "lcms2.dll" (ByVal ptrToCmsLogErrorHandlerFunction As Long)

'Profile create/release functions; white points declared as ByVal Longs can typically be set to NULL to use the default D50 value
Private Declare Function cmsCloseProfile Lib "lcms2.dll" (ByVal srcProfile As Long) As Long
Private Declare Function cmsCreateBCHSWabstractProfile Lib "lcms2.dll" (ByVal nLUTPoints As Long, ByVal newBrightness As Double, ByVal newContrast As Double, ByVal newHue As Double, ByVal newSaturation As Double, ByVal srcTemp As Long, ByVal dstTemp As Long) As Long
Private Declare Function cmsCreateGrayProfile Lib "lcms2.dll" (ByVal ptrToWhitePointxyY As Long, ByVal sourceToneCurve As Long) As Long
Private Declare Function cmsCreateLab2Profile Lib "lcms2.dll" (ByVal ptrToWhitePointxyY As Long) As Long
Private Declare Function cmsCreateLab4Profile Lib "lcms2.dll" (ByVal ptrToWhitePointxyY As Long) As Long
Private Declare Function cmsCreate_sRGBProfile Lib "lcms2.dll" () As Long
Private Declare Function cmsCreateRGBProfile Lib "lcms2.dll" (ByVal ptrToWhitePointxyY As Long, ByVal ptrTo3xyYPrimaries As Long, ByVal ptrTo3ToneCurves As Long) As Long
Private Declare Function cmsCreateXYZProfile Lib "lcms2.dll" () As Long
Private Declare Function cmsOpenProfileFromMem Lib "lcms2.dll" (ByVal ptrProfile As Long, ByVal profileSizeInBytes As Long) As Long
Private Declare Function cmsSaveProfileToMem Lib "lcms2.dll" (ByVal srcProfile As Long, ByVal dstPtr As Long, ByRef sizeRequiredInBytes As Long) As Long
 
'Profile information functions
Private Declare Function cmsGetHeaderRenderingIntent Lib "lcms2.dll" (ByVal hProfile As Long) As LCMS_RENDERING_INTENT

'Tone curve creation/destruction
Private Declare Function cmsBuildParametricToneCurve Lib "lcms2.dll" (ByVal ContextID As Long, ByVal tcType As Long, ByVal ptrToFirstParam As Long) As Long
Private Declare Function cmsBuildGamma Lib "lcms2.dll" (ByVal ContextID As Long, ByVal gammaValue As Double) As Long
Private Declare Sub cmsFreeToneCurve Lib "lcms2.dll" (ByVal srcToneCurve As Long)

'Transform functions
Private Declare Function cmsCreateTransform Lib "lcms2.dll" (ByVal hInputProfile As Long, ByVal hInputFormat As LCMS_PIXEL_FORMAT, ByVal hOutputProfile As Long, ByVal hOutputFormat As LCMS_PIXEL_FORMAT, ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT, ByVal trnsFlags As LCMS_TRANSFORM_FLAGS) As Long
Private Declare Function cmsCreateMultiprofileTransform Lib "lcms2.dll" (ByVal ptrToFirstProfile As Long, ByVal numOfProfiles As Long, ByVal hInputFormat As LCMS_PIXEL_FORMAT, ByVal hOutputFormat As LCMS_PIXEL_FORMAT, ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT, ByVal trnsFlags As LCMS_TRANSFORM_FLAGS) As Long
Private Declare Sub cmsDeleteTransform Lib "lcms2.dll" (ByVal hTransform As Long)

'Color space conversions; any conversion that requires an XYZ WhitePoint can pass null for default D50 values
Private Declare Sub cmsLab2XYZ Lib "lcms2.dll" (ByVal ptrToWhitePointXYZ As Long, ByRef dstXYZ As LCMS_XYZ, ByRef srcLab As LCMS_Lab)
Private Declare Sub cmsXYZ2Lab Lib "lcms2.dll" (ByVal ptrToWhitePointXYZ As Long, ByRef dstLab As LCMS_Lab, ByRef srcXYZ As LCMS_XYZ)
Private Declare Sub cmsXYZ2xyY Lib "lcms2.dll" (ByRef dstxyY As LCMS_xyY, ByRef srcXYZ As LCMS_XYZ)
Private Declare Sub cmsxyY2XYZ Lib "lcms2.dll" (ByRef dstXYZ As LCMS_XYZ, ByRef srcxyY As LCMS_xyY)
Private Declare Function cmsWhitePointFromTemp Lib "lcms2.dll" (ByRef dstWhitePointxyY As LCMS_xyY, ByVal srcTemperature As Double) As Long
Private Declare Function cmsTempFromWhitePoint Lib "lcms2.dll" (ByRef dstTemperature As Double, ByRef srcWhitePointxyY As LCMS_xyY) As Long

'Pointers to the constant XYZ/xyY declarations for D50
Private Declare Function cmsD50_XYZ Lib "lcms2.dll" () As Long
Private Declare Function cmsD50_xyY Lib "lcms2.dll" () As Long

'Similar internal functions for D65 (which is used by a number of RGB spaces, e.g. Adobe and sRGB)
Private m_D65_XYZ() As Double, m_D65_xyY() As Double

'Actual transform application functions
Private Declare Sub cmsDoTransform Lib "lcms2.dll" (ByVal hTransform As Long, ByVal ptrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal numOfPixelsToTransform As Long)

'In 2.8, a dedicated line/stride transform function was added to LittleCMS.  Here is what the documentation says:
' "This function translates bitmaps with complex organization. Each bitmap may contain several lines, and every line
'  may have padding. The distance from one line to the next one is BytesPerLine{In/Out}.  In planar formats, each line
'  may hold several planes, and each plane may have padding. Padding of lines and planes should be same across the whole
'  bitmap, i.e. all lines in a bitmap must be padded the same way. This function may be more efficient that repeated calls
'  to cmsDoTransform(), especially when customized plug-ins are being used."
'
'I do not currently make use of this function, but given the efficiency caveat above, it may be worth investigating in the future.
Private Declare Sub cmsDoTransformLineStride Lib "lcms2.dll" (ByVal hTransform As Long, ByVal ptrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal numOfPixelsPerLine As Long, ByVal numOfLines As Long, ByVal bytesPerLineIn As Long, ByVal bytesPerLineOut As Long, ByVal bytesPerPlaneIn As Long, ByVal bytesPerPlaneOut As Long)

'A single LittleCMS handle is maintained for the life of a PD instance; see InitializeLCMS and ReleaseLCMS, below.
Private m_LCMSHandle As Long

'LittleCMS requires a CDECL callback for its error handler function.  VB can't provide this, so we use a workaround provided by LaVolpe.
Private m_CdeclWorkaround As cUniversalDLLCalls, m_CdeclCallback As Long

'Initialize LittleCMS.  Do not call this until you have verified the LCMS plugin's existence
' (typically via the PluginManager module)
Public Function InitializeLCMS() As Boolean
    
    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    Dim lcmsPath As String
    lcmsPath = g_PluginPath & "lcms2.dll"
    m_LCMSHandle = LoadLibrary(StrPtr(lcmsPath))
    InitializeLCMS = CBool(m_LCMSHandle <> 0)
    
    #If DEBUGMODE = 1 Then
        
        'Set up an error logger.  Note that this WILL CRASH THAT PROGRAM after a log due to StdCall behavior.  As such,
        ' it's only good for retrieving a single error (before everything goes to shit).
        If InitializeLCMS Then
            Set m_CdeclWorkaround = New cUniversalDLLCalls
            m_CdeclCallback = m_CdeclWorkaround.ThunkFor_CDeclCallbackToVB(AddressOf cmsErrorHandler, 3)
            Call cmsSetLogErrorHandler(m_CdeclCallback)
            pdDebug.LogAction "LittleCMS callback successfully set: " & m_CdeclCallback
        End If
        
        If (Not InitializeLCMS) Then
            pdDebug.LogAction "WARNING!  LoadLibrary failed to load LittleCMS.  Last DLL error: " & Err.LastDllError
            pdDebug.LogAction "(FYI, the attempted path was: " & lcmsPath & ")"
        End If
    #End If
    
    'Initialize D65 primaries as well; these are helpful shortcuts when assembling RGB profiles on-the-fly
    ReDim m_D65_XYZ(0 To 2) As Double
    ReDim m_D65_xyY(0 To 2) As Double
    m_D65_XYZ(0) = 0.95045471
    m_D65_XYZ(1) = 1#
    m_D65_XYZ(2) = 1.08905029
    m_D65_xyY(0) = 0.3127
    m_D65_xyY(1) = 0.329
    m_D65_xyY(2) = 1#
    
End Function

'When PD closes, make sure to release our library handle
Public Sub ReleaseLCMS()
    If (m_CdeclCallback <> 0) Then
        m_CdeclWorkaround.ThunkRelease_CDECL m_CdeclCallback
        m_CdeclCallback = 0
        Set m_CdeclWorkaround = Nothing
    End If
    If (m_LCMSHandle <> 0) Then FreeLibrary m_LCMSHandle
    g_LCMSEnabled = False
End Sub

'After LittleCMS has been initialized, you can call this function to retrieve its current version.
' The version will always be formatted as "Major.Minor.0.0".
Public Function GetLCMSVersion() As String
    
    Dim versionAsLong As Long
    versionAsLong = cmsGetEncodedCMMversion()
    
    'Split the version by zeroes
    Dim versionAsString() As String
    versionAsString = Split(CStr(versionAsLong), "0", , vbBinaryCompare)
    
    If VB_Hacks.IsArrayInitialized(versionAsString) Then
        If (UBound(versionAsString) >= 1) Then
            GetLCMSVersion = versionAsString(0) & "." & versionAsString(1) & ".0.0"
        Else
            GetLCMSVersion = "0.0.0.0"
        End If
    Else
        GetLCMSVersion = "0.0.0.0"
    End If
    
End Function

'Fake wrappers to emulate the cmsD50 functions provided by LittleCMS
Private Function cmsD65_XYZ() As Long
    cmsD65_XYZ = VarPtr(m_D65_XYZ(0))
End Function

Private Function cmsD65_xyY() As Long
    cmsD65_xyY = VarPtr(m_D65_xyY(0))
End Function

Public Function LCMS_GetIlluminantxyY(ByRef dstxyY As LCMS_xyY, Optional ByVal srcIlluminant As LCMS_ILLUMINANT = cmsILLUMINANT_TYPE_D50) As Boolean
    If (srcIlluminant >= cmsILLUMINANT_TYPE_D50) And (srcIlluminant <= cmsILLUMINANT_TYPE_F8) Then
        Dim srcTemperature As Double
        LCMS_GetIlluminantTemperature srcTemperature, srcIlluminant
        LCMS_GetIlluminantxyY = CBool(cmsWhitePointFromTemp(dstxyY, srcTemperature) <> 0)
    Else
        LCMS_GetIlluminantxyY = False
    End If
End Function

Public Function LCMS_GetIlluminantTemperature(ByRef dstTemperature As Double, Optional ByVal srcIlluminant As LCMS_ILLUMINANT = cmsILLUMINANT_TYPE_D50) As Boolean
    
    Select Case srcIlluminant
        Case cmsILLUMINANT_TYPE_UNKNOWN
            LCMS_GetIlluminantTemperature = False
        Case cmsILLUMINANT_TYPE_D50
            LCMS_GetIlluminantTemperature = True
            dstTemperature = 5003
        Case cmsILLUMINANT_TYPE_D65
            LCMS_GetIlluminantTemperature = True
            dstTemperature = 6504
        Case cmsILLUMINANT_TYPE_D93
            LCMS_GetIlluminantTemperature = True
            dstTemperature = 9303
        Case cmsILLUMINANT_TYPE_F2
            LCMS_GetIlluminantTemperature = True
            dstTemperature = 4230
        Case cmsILLUMINANT_TYPE_D55
            LCMS_GetIlluminantTemperature = True
            dstTemperature = 5503
        Case cmsILLUMINANT_TYPE_A
            LCMS_GetIlluminantTemperature = True
            dstTemperature = 2856
        Case cmsILLUMINANT_TYPE_E
            LCMS_GetIlluminantTemperature = True
            dstTemperature = 5454
        Case cmsILLUMINANT_TYPE_F8
            LCMS_GetIlluminantTemperature = True
            dstTemperature = 5000
        Case Else
            LCMS_GetIlluminantTemperature = False
    End Select

End Function

Public Function LCMS_CreateTwoProfileTransform(ByVal hInputProfile As Long, ByVal hOutputProfile As Long, Optional ByVal hInputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal hOutputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = cmsFLAGS_COPY_ALPHA) As Long
    LCMS_CreateTwoProfileTransform = cmsCreateTransform(hInputProfile, hInputFormat, hOutputProfile, hOutputFormat, trnsRenderingIntent, trnsFlags)
End Function

Public Function LCMS_CreateMultiProfileTransform(ByRef hProfiles() As Long, ByVal numOfProfiles As Long, Optional ByVal hInputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal hOutputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = cmsFLAGS_COPY_ALPHA) As Long
    LCMS_CreateMultiProfileTransform = cmsCreateMultiprofileTransform(VarPtr(hProfiles(0)), numOfProfiles, hInputFormat, hOutputFormat, trnsRenderingIntent, trnsFlags)
End Function

Public Function LCMS_CreateInPlaceTransformForDIB(ByVal hInputProfile As Long, ByVal hOutputProfile As Long, ByRef srcDIB As pdDIB, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = cmsFLAGS_COPY_ALPHA) As Long
    
    Dim pxFormat As LCMS_PIXEL_FORMAT
    If (srcDIB.GetDIBColorDepth = 32) Then
        pxFormat = TYPE_BGRA_8
    Else
        pxFormat = TYPE_BGR_8
    End If
    
    LCMS_CreateInPlaceTransformForDIB = cmsCreateTransform(hInputProfile, pxFormat, hOutputProfile, pxFormat, trnsRenderingIntent, trnsFlags)
    
End Function

Public Function LCMS_DeleteTransform(ByRef hTransform As Long) As Boolean
    cmsDeleteTransform hTransform
    hTransform = 0
    LCMS_DeleteTransform = True
End Function

Public Function LCMS_GetProfileRenderingIntent(ByVal hProfile As Long) As LCMS_RENDERING_INTENT
    LCMS_GetProfileRenderingIntent = cmsGetHeaderRenderingIntent(hProfile)
End Function

Public Function LCMS_LoadProfileFromMemory(ByVal ptrToProfile As Long, ByVal sizeOfProfileInBytes As Long) As Long
    LCMS_LoadProfileFromMemory = cmsOpenProfileFromMem(ptrToProfile, sizeOfProfileInBytes)
End Function

'Little CMS has its own "load from file" function, but it isn't Unicode-aware, so we just slam the file into a byte array
' and use the "load from memory" function instead.
Public Function LCMS_LoadProfileFromFile(ByVal profilePath As String) As Long
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO

    'Start by loading the specified path into a byte array
    Dim tmpProfileArray() As Byte
        
    If cFile.FileExist(profilePath) Then
        
        If (Not cFile.LoadFileAsByteArray(profilePath, tmpProfileArray)) Then
            LCMS_LoadProfileFromFile = 0
            Exit Function
        End If
        
    Else
        LCMS_LoadProfileFromFile = 0
        Exit Function
    End If
    
    LCMS_LoadProfileFromFile = cmsOpenProfileFromMem(VarPtr(tmpProfileArray(0)), UBound(tmpProfileArray) + 1)
    
End Function

Public Function LCMS_LoadStockGrayProfile() As Long
    Dim tmpToneCurve As Long
    tmpToneCurve = LCMS_GetBasicToneCurve(1#)
    LCMS_LoadStockGrayProfile = cmsCreateGrayProfile(cmsD50_xyY(), tmpToneCurve)
    LittleCMS.LCMS_FreeToneCurve tmpToneCurve
End Function

'Linear RGB profile is identical to sRGB except for the gamma curve, which is (obviously) flat
Public Function LCMS_LoadLinearRGBProfile() As Long
    
    Dim rgbPrimaries() As Double
    ReDim rgbPrimaries(0 To 8) As Double
    rgbPrimaries(0) = 0.639998686
    rgbPrimaries(1) = 0.330010138
    rgbPrimaries(2) = 1#
    
    rgbPrimaries(3) = 0.300003784
    rgbPrimaries(4) = 0.600003357
    rgbPrimaries(5) = 1#
    
    rgbPrimaries(6) = 0.150002046
    rgbPrimaries(7) = 0.059997204
    rgbPrimaries(8) = 1#
    
    Dim rgbToneCurves() As Long
    ReDim rgbToneCurves(0 To 2) As Long
    rgbToneCurves(0) = LCMS_GetBasicToneCurve(1#)
    rgbToneCurves(1) = rgbToneCurves(0)
    rgbToneCurves(2) = rgbToneCurves(0)
    
    LCMS_LoadLinearRGBProfile = cmsCreateRGBProfile(cmsD65_xyY(), VarPtr(rgbPrimaries(0)), VarPtr(rgbToneCurves(0)))
    
    'The intermediate tone curve *must* be freed now, as it's never directly exposed to the caller
    LCMS_FreeToneCurve rgbToneCurves(0)
    
End Function

Public Function LCMS_LoadStockSRGBProfile() As Long
    LCMS_LoadStockSRGBProfile = cmsCreate_sRGBProfile()
End Function

Public Function LCMS_LoadStockLabProfile(Optional ByVal useVersion4 As Boolean = True) As Long
    If useVersion4 Then
        LCMS_LoadStockLabProfile = cmsCreateLab4Profile(0&)
    Else
        LCMS_LoadStockLabProfile = cmsCreateLab2Profile(0&)
    End If
End Function

Public Function LCMS_SaveProfileToArray(ByVal hProfile As Long, ByRef dstArray() As Byte) As Boolean
    
    Dim profSize As Long
    
    'Passing a null pointer will fill the "profile size" parameter with the required destination size
    If (cmsSaveProfileToMem(hProfile, 0, profSize) <> 0) Then
        ReDim dstArray(0 To profSize - 1) As Byte
        LCMS_SaveProfileToArray = CBool(cmsSaveProfileToMem(hProfile, VarPtr(dstArray(0)), profSize) <> 0)
    Else
        LCMS_SaveProfileToArray = False
    End If
    
End Function

Public Function LCMS_CreateAbstractBCHSProfile(Optional ByVal newBrightness As Double = 0#, Optional ByVal newContrast As Double = 1#, Optional ByVal newHue As Double = 0#, Optional ByVal newSaturation As Double = 0#, Optional ByVal srcTemp As Long = 0, Optional ByVal dstTemp As Long = 0) As Long
    LCMS_CreateAbstractBCHSProfile = cmsCreateBCHSWabstractProfile(16, newBrightness, newContrast, newHue, newSaturation, srcTemp, dstTemp)
End Function

Public Function LCMS_CloseProfileHandle(ByRef srcHandle As Long) As Boolean
    LCMS_CloseProfileHandle = CBool(cmsCloseProfile(srcHandle) <> 0)
    If LCMS_CloseProfileHandle Then srcHandle = 0
End Function

Private Function LCMS_GetBasicToneCurve(Optional ByVal srcGamma As Double = 1#) As Long
    LCMS_GetBasicToneCurve = cmsBuildGamma(0&, srcGamma)
End Function

Public Function LCMS_FreeToneCurve(ByRef hCurve As Long) As Boolean
    cmsFreeToneCurve hCurve
    hCurve = 0
    LCMS_FreeToneCurve = True
End Function

'Apply an already-created transform to a pdDIB object.
Public Function LCMS_ApplyTransformToDIB(ByRef srcDIB As pdDIB, ByVal hTransform As Long) As Boolean
    
    If (Not (srcDIB Is Nothing)) And (hTransform <> 0) Then
        
        '32-bpp DIBs can be applied in one fell swoop, since there are no scanline padding issues
        If (srcDIB.GetDIBColorDepth = 32) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Applying ICC transform to 32-bpp DIB..."
            #End If
            
            cmsDoTransform hTransform, srcDIB.GetDIBPointer, srcDIB.GetDIBPointer, srcDIB.GetDIBWidth * srcDIB.GetDIBHeight
                    
        '24-bpp DIBs may have scanline padding issues.  We must process them one line at a time.
        Else
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Applying ICC transform to 24-bpp DIB..."
            #End If
            
            Dim i As Long, iWidth As Long, iScanWidth As Long, iScanStart As Long
            iWidth = srcDIB.GetDIBWidth
            iScanStart = srcDIB.GetDIBPointer
            iScanWidth = srcDIB.GetDIBStride
            
            For i = 0 To srcDIB.GetDIBHeight - 1
                cmsDoTransform hTransform, iScanStart + i * iScanWidth, iScanStart + i * iScanWidth, iWidth
            Next i
        
        End If
        
        'The "cmsDoTransform" function has no return, so we assume success if passed a valid DIB and transform
        LCMS_ApplyTransformToDIB = True
        
    End If
        
End Function

Public Sub LCMS_TransformArbitraryMemory(ByVal srcPointer As Long, ByVal dstPointer As Long, ByVal WidthInPixels As Long, ByVal hTransform As Long)
    cmsDoTransform hTransform, srcPointer, dstPointer, WidthInPixels
End Sub

'Given a target DIB with a valid .ICCProfile object, apply said profile to said DIB.
' (NOTE!  If the source image is 32-bpp, with premultiplied alpha, you need to unpremultiply alpha prior to
'         calling this function; otherwise, the end result will be invalid.)
Public Function ApplyICCProfileToPDDIB(ByRef targetDIB As pdDIB) As Boolean
    
    ApplyICCProfileToPDDIB = False
    
    If (targetDIB Is Nothing) Then
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  LittleCMS.ApplyICCProfileToPDDIB was passed a null pdDIB."
        #End If
        Exit Function
    End If
    
    'Before doing anything else, make sure we actually have an ICC profile to apply!
    If (Not targetDIB.ICCProfile.HasICCData) Then
        Message "ICC transform requested, but no data found.  Abandoning attempt."
        Exit Function
    End If
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Using embedded ICC profile to convert image to sRGB space for editing..."
    #End If
    
    'Start by creating two LCMS profile handles:
    ' 1) a source profile (the in-memory copy of the ICC profile associated with this DIB)
    ' 2) a destination profile (the current PhotoDemon working space)
    Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile
    Set srcProfile = New pdLCMSProfile
    Set dstProfile = New pdLCMSProfile
    
    If srcProfile.CreateFromPDDib(targetDIB) Then
        
        If dstProfile.CreateSRGBProfile() Then
            
            'DISCLAIMER! Until rendering intent has a dedicated preference, PD defaults to perceptual render intent.
            ' This provides better results on most images, it correctly preserves gamut, and it is the standard
            ' behavior for PostScript workflows.  See http://fieryforums.efi.com/showthread.php/835-Rendering-Intent-Control-for-Embedded-Profiles
            ' Also see: https://developer.mozilla.org/en-US/docs/ICC_color_correction_in_Firefox)
            '
            'For future reference, I've left the code below for retrieving rendering intent from the source profile
            Dim targetRenderingIntent As LCMS_RENDERING_INTENT
            targetRenderingIntent = INTENT_PERCEPTUAL
            'targetRenderingIntent = srcProfile.GetRenderingIntent
            
            'Create a transform that uses the target DIB as both the source and destination
            Dim cTransform As pdLCMSTransform
            Set cTransform = New pdLCMSTransform
            If cTransform.CreateInPlaceTransformForDIB(targetDIB, srcProfile, dstProfile, targetRenderingIntent) Then
                
                'LittleCMS 2.0 allows us to free our source profiles immediately after a transform is created.
                ' (Note that we don't *need* to do this, nor does this code leak if we don't manually free both
                '  profiles, but as we're about to do an energy- and memory-intensive operation, it doesn't
                '  hurt to free the profiles now.)
                Set srcProfile = Nothing: Set dstProfile = Nothing
                
                If cTransform.ApplyTransformToPDDib(targetDIB) Then
                
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "ICC profile transformation successful.  Image now lives in the current RGB working space."
                    #End If
                    
                    targetDIB.ICCProfile.MarkSuccessfulProfileApplication
                    ApplyICCProfileToPDDIB = True
                    
                End If
                
                'Note that we could free the transform here, but it's unnecessary.  (The pdLCMSTransform class
                ' is self-freeing upon destruction.)
                
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  LittleCMS.ApplyICCProfileToPDDIB failed to create a valid transformation handle!"
                #End If
            End If
        
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  LittleCMS.ApplyICCProfileToPDDib failed to create a valid destination profile handle."
            #End If
        End If
    
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  LittleCMS.ApplyICCProfileToPDDib failed to create a valid source profile handle."
        #End If
    End If
    
End Function

Private Function cmsErrorHandler(ByVal ContextID As Long, ByVal cmsError As Long, ByVal ptrToText As Long) As Long
    #If DEBUGMODE = 1 Then
        Dim cUnicode As pdUnicode
        Set cUnicode = New pdUnicode
        
        Dim errorMsg As String
        errorMsg = cUnicode.ConvertCharPointerToVBString(ptrToText, False)
        
        pdDebug.LogAction "WARNING!  LittleCMS error occurred (#" & cmsError & "): " & errorMsg
    #End If
End Function
