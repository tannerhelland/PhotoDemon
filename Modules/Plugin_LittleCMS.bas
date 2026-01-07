Attribute VB_Name = "LittleCMS"
'***************************************************************************
'LittleCMS Interface
'Copyright 2016-2026 by Tanner Helland
'Created: 21/April/16
'Last updated: 28/August/25
'Last update: add wrapper for creating custom RGB profiles via custom tone curves
'
'Module for handling all LittleCMS interfacing.  This module is pointless without the accompanying
' LittleCMS plugin, which will be in the App/PhotoDemon/Plugins subdirectory as "lcms2.dll".
'
'LittleCMS is a free, open-source color management library.  You can learn more about it here:
' http://www.littlecms.com/
'
'PhotoDemon has been designed against v2.8.0.  It may not work with other versions.
' Additional documentation regarding the use of LittleCMS is available as part of the official
' LittleCMS library, available from https://github.com/mm2/Little-CMS.
'
'LittleCMS is available under the MIT license.  Please see the App/PhotoDemon/Plugins/lcms2-LICENSE.txt
' file for questions regarding copyright or licensing.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Type LCMS_xyY
    x As Double
    y As Double
    YY As Double
End Type

'Not all color formats are used at present
'
'Public Type LCMS_Lab
'    l As Double
'    a As Double
'    b As Double
'End Type
'
'Public Type LCMS_LCh
'    l As Double
'    c As Double
'    h As Double
'End Type

'Public Type LCMS_JCh
'    j As Double
'    c As Double
'    h As Double
'End Type
'
'Public Type LCMS_XYZ
'    x As Double
'    y As Double
'    z As Double
'End Type

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
    
    '(Added by Tanner, but it turns out 32-bit integers are not supported!)
    'TYPE_RGB_32 = &H4001C
    'TYPE_RGB_32_PLANAR = &H4101C
    
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
    
    '(Added by Tanner, but it turns out 32-bit integers are not supported!)
    'TYPE_RGBA_32 = &H4009C
    'TYPE_RGBA_32_PLANAR = &H4109C
    
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
    
    '(Added by Tanner)
    TYPE_GRAYA_HALF_FLT = &H43008A
    TYPE_GRAYA_FLT = &H43008C
    TYPE_GRAYA_DBL = &H430088
    TYPE_GRAYA_HALF_FLT_PLANAR = &H43108A
    TYPE_GRAYA_FLT_PLANAR = &H43108C
    TYPE_GRAYA_DBL_PLANAR = &H431088
    
    TYPE_GRAY_HALF_FLT = &H43000A
    TYPE_RGB_HALF_FLT = &H44001A
    TYPE_RGBA_HALF_FLT = &H44009A
    TYPE_CMYK_HALF_FLT = &H460022

    TYPE_ARGB_HALF_FLT = &H44409A
    TYPE_BGR_HALF_FLT = &H44041A
    TYPE_BGRA_HALF_FLT = &H44449A
    TYPE_ABGR_HALF_FLT = &H44041A
    
    '(Added by Tanner)
    TYPE_RGB_HALF_FLT_PLANAR = &H44101A
    TYPE_RGBA_HALF_FLT_PLANAR = &H44109A
    TYPE_RGB_FLT_PLANAR = &H44101C
    TYPE_RGBA_FLT_PLANAR = &H44109C
    TYPE_RGB_DBL_PLANAR = &H441018
    TYPE_RGBA_DBL_PLANAR = &H441098
    TYPE_RGBA_DBL = &H440098
    
    'These flags are *not* automatically defined by LCMS; I've defined them to allow for OR'ing with
    ' existing constants (since VB makes bit-shifting such a PITA)
    FLAG_ALPHAPRESENT = &H80&
    FLAG_MINISWHITE = &H2000&
    FLAG_PLANAR = &H1000&
    FLAG_SE = &H800&
    
End Enum

#If False Then
    Private Const TYPE_GRAY_8 = &H30009, TYPE_GRAY_8_REV = &H32009, TYPE_GRAY_16 = &H3000A, TYPE_GRAY_16_REV = &H3200A, TYPE_GRAY_16_SE = &H3080A, TYPE_GRAYA_8 = &H30089, TYPE_GRAYA_16 = &H3008A, TYPE_GRAYA_16_SE = &H3088A, TYPE_GRAYA_8_PLANAR = &H31089, TYPE_GRAYA_16_PLANAR = &H3108A, TYPE_RGB_8 = &H40019, TYPE_RGB_8_PLANAR = &H41019, TYPE_BGR_8 = &H40419, TYPE_BGR_8_PLANAR = &H41419, TYPE_RGB_16 = &H4001A, TYPE_RGB_16_PLANAR = &H4101A
    Private Const TYPE_RGB_16_SE = &H4081A, TYPE_BGR_16 = &H4041A, TYPE_BGR_16_PLANAR = &H4141A, TYPE_BGR_16_SE = &H40C1A, TYPE_RGBA_8 = &H40099, TYPE_RGBA_8_PLANAR = &H41099, TYPE_ARGB_8_PLANAR = &H45099, TYPE_ABGR_8_PLANAR = &H41499, TYPE_BGRA_8_PLANAR = &H45499, TYPE_RGBA_16 = &H4009A, TYPE_RGBA_16_PLANAR = &H4109A, TYPE_RGBA_16_SE = &H4089A, TYPE_ARGB_8 = &H44099, TYPE_ARGB_16 = &H4409A, TYPE_ABGR_8 = &H40499, TYPE_ABGR_16 = &H4049A
    Private Const TYPE_ABGR_16_PLANAR = &H4149A, TYPE_ABGR_16_SE = &H40C9A, TYPE_BGRA_8 = &H44499, TYPE_BGRA_16 = &H4449A, TYPE_BGRA_16_SE = &H4489A, TYPE_CMY_8 = &H50019, TYPE_CMY_8_PLANAR = &H51019, TYPE_CMY_16 = &H5001A, TYPE_CMY_16_PLANAR = &H5101A, TYPE_CMY_16_SE = &H5081A, TYPE_CMYK_8 = &H60021, TYPE_CMYKA_8 = &H600A1, TYPE_CMYK_8_REV = &H62021, TYPE_YUVK_8 = &H62021, TYPE_CMYK_8_PLANAR = &H61021, TYPE_CMYK_16 = &H60022
    Private Const TYPE_CMYK_16_REV = &H62022, TYPE_YUVK_16 = &H62022, TYPE_CMYK_16_PLANAR = &H61022, TYPE_CMYK_16_SE = &H60822, TYPE_KYMC_8 = &H60421, TYPE_KYMC_16 = &H60422, TYPE_KYMC_16_SE = &H60C22, TYPE_KCMY_8 = &H64021, TYPE_KCMY_8_REV = &H66021, TYPE_KCMY_16 = &H64022, TYPE_KCMY_16_REV = &H66022, TYPE_KCMY_16_SE = &H64822, TYPE_CMYK5_8 = &H130029, TYPE_CMYK5_16 = &H13002A, TYPE_CMYK5_16_SE = &H13082A, TYPE_KYMC5_8 = &H130429
    Private Const TYPE_KYMC5_16 = &H13042A, TYPE_KYMC5_16_SE = &H130C2A, TYPE_CMYK6_8 = &H140031, TYPE_CMYK6_8_PLANAR = &H141031, TYPE_CMYK6_16 = &H140032, TYPE_CMYK6_16_PLANAR = &H141032, TYPE_CMYK6_16_SE = &H140832, TYPE_CMYK7_8 = &H150039, TYPE_CMYK7_16 = &H15003A, TYPE_CMYK7_16_SE = &H15083A, TYPE_KYMC7_8 = &H150439, TYPE_KYMC7_16 = &H15043A, TYPE_KYMC7_16_SE = &H150C3A, TYPE_CMYK8_8 = &H160041, TYPE_CMYK8_16 = &H160042
    Private Const TYPE_CMYK8_16_SE = &H160842, TYPE_KYMC8_8 = &H160441, TYPE_KYMC8_16 = &H160442, TYPE_KYMC8_16_SE = &H160C42, TYPE_CMYK9_8 = &H170049, TYPE_CMYK9_16 = &H17004A, TYPE_CMYK9_16_SE = &H17084A, TYPE_KYMC9_8 = &H170449, TYPE_KYMC9_16 = &H17044A, TYPE_KYMC9_16_SE = &H170C4A, TYPE_CMYK10_8 = &H180051, TYPE_CMYK10_16 = &H180052, TYPE_CMYK10_16_SE = &H180852, TYPE_KYMC10_8 = &H180451, TYPE_KYMC10_16 = &H180452
    Private Const TYPE_KYMC10_16_SE = &H180C52, TYPE_CMYK11_8 = &H190059, TYPE_CMYK11_16 = &H19005A, TYPE_CMYK11_16_SE = &H19085A, TYPE_KYMC11_8 = &H190459, TYPE_KYMC11_16 = &H19045A, TYPE_KYMC11_16_SE = &H190C5A, TYPE_CMYK12_8 = &H1A0061, TYPE_CMYK12_16 = &H1A0062, TYPE_CMYK12_16_SE = &H1A0862, TYPE_KYMC12_8 = &H1A0461, TYPE_KYMC12_16 = &H1A0462, TYPE_KYMC12_16_SE = &H1A0C62, TYPE_XYZ_16 = &H9001A, TYPE_Lab_8 = &HA0019
    Private Const TYPE_ALab_8 = &HA0499, TYPE_Lab_16 = &HA001A, TYPE_Yxy_16 = &HE001A, TYPE_YCbCr_8 = &H70019, TYPE_YCbCr_8_PLANAR = &H71019, TYPE_YCbCr_16 = &H7001A, TYPE_YCbCr_16_PLANAR = &H7101A, TYPE_YCbCr_16_SE = &H7081A, TYPE_YUV_8 = &H80019, TYPE_YUV_8_PLANAR = &H81019, TYPE_YUV_16 = &H8001A, TYPE_YUV_16_PLANAR = &H8101A, TYPE_YUV_16_SE = &H8081A, TYPE_HLS_8 = &HD0019, TYPE_HLS_8_PLANAR = &HD1019, TYPE_HLS_16 = &HD001A
    Private Const TYPE_HLS_16_PLANAR = &HD101A, TYPE_HLS_16_SE = &HD081A, TYPE_HSV_8 = &HC0019, TYPE_HSV_8_PLANAR = &HC1019, TYPE_HSV_16 = &HC001A, TYPE_HSV_16_PLANAR = &HC101A, TYPE_HSV_16_SE = &HC081A, TYPE_NAMED_COLOR_INDEX = &HA&, TYPE_XYZ_FLT = &H49001C, TYPE_Lab_FLT = &H4A001C, TYPE_GRAY_FLT = &H43000C, TYPE_RGB_FLT = &H44001C, TYPE_CMYK_FLT = &H460024, TYPE_XYZA_FLT = &H49009C, TYPE_LabA_FLT = &H4A009C, TYPE_RGBA_FLT = &H44009C
    Private Const TYPE_XYZ_DBL = &H490018, TYPE_Lab_DBL = &H4A0018, TYPE_GRAY_DBL = &H430008, TYPE_RGB_DBL = &H440018, TYPE_CMYK_DBL = &H460020, TYPE_LabV2_8 = &H1E0019, TYPE_ALabV2_8 = &H1E0499, TYPE_LabV2_16 = &H1E001A, TYPE_GRAY_HALF_FLT = &H43000A, TYPE_RGB_HALF_FLT = &H44001A, TYPE_RGBA_HALF_FLT = &H44009A, TYPE_CMYK_HALF_FLT = &H460022, TYPE_ARGB_HALF_FLT = &H44409A, TYPE_BGR_HALF_FLT = &H44041A, TYPE_BGRA_HALF_FLT = &H44449A, TYPE_ABGR_HALF_FLT = &H44041A
    Private Const FLAG_ALPHAPRESENT = &H80&, FLAG_MINISWHITE = &H2000&, FLAG_PLANAR = &H1000&, FLAG_SE = &H800&
#End If

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

#If False Then
    Private Const INTENT_PERCEPTUAL = 0, INTENT_RELATIVE_COLORIMETRIC = 1, INTENT_SATURATION = 2, INTENT_ABSOLUTE_COLORIMETRIC = 3, INTENT_PRESERVE_K_ONLY_PERCEPTUAL = 10, INTENT_PRESERVE_K_ONLY_RELATIVE_COLORIMETRIC = 11, INTENT_PRESERVE_K_ONLY_SATURATION = 12, INTENT_PRESERVE_K_PLANE_PERCEPTUAL = 13, INTENT_PRESERVE_K_PLANE_RELATIVE_COLORIMETRIC = 14, INTENT_PRESERVE_K_PLANE_SATURATION = 15
#End If

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

#If False Then
    Private Const cmsFLAGS_NOCACHE = &H40&, cmsFLAGS_NOOPTIMIZE = &H100&, cmsFLAGS_NULLTRANSFORM = &H200&, cmsFLAGS_GAMUTCHECK = &H1000&, cmsFLAGS_SOFTPROOFING = &H4000&, cmsFLAGS_BLACKPOINTCOMPENSATION = &H2000&, cmsFLAGS_NOWHITEONWHITEFIXUP = &H4&, cmsFLAGS_HIGHRESPRECALC = &H400&, cmsFLAGS_LOWRESPRECALC = &H800&, cmsFLAGS_8BITS_DEVICELINK = &H8&, cmsFLAGS_GUESSDEVICECLASS = &H20&, cmsFLAGS_KEEP_SEQUENCE = &H80&, cmsFLAGS_FORCE_CLUT = &H2&, cmsFLAGS_CLUT_POST_LINEARIZATION = &H1&, cmsFLAGS_CLUT_PRE_LINEARIZATION = &H10&, cmsFLAGS_NODEFAULTRESOURCEDEF = &H1000000, cmsFLAGS_COPY_ALPHA = &H4000000
#End If

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

#If False Then
    Private Const cmsILLUMINANT_TYPE_UNKNOWN = &H0, cmsILLUMINANT_TYPE_D50 = &H1, cmsILLUMINANT_TYPE_D65 = &H2, cmsILLUMINANT_TYPE_D93 = &H3, cmsILLUMINANT_TYPE_F2 = &H4, cmsILLUMINANT_TYPE_D55 = &H5, cmsILLUMINANT_TYPE_A = &H6, cmsILLUMINANT_TYPE_E = &H7, cmsILLUMINANT_TYPE_F8 = &H8
#End If

'Want to pull basic information from an ICC profile?  These "quick" enums can be retrieved, and they are all
' Unicode-aware.
Public Enum LCMS_INFOTYPE
    cmsInfoDescription = 0
    cmsInfoManufacturer = 1
    cmsInfoModel = 2
    cmsInfoCopyright = 3
End Enum

#If False Then
    Private Const cmsInfoDescription = 0, cmsInfoManufacturer = 1, cmsInfoModel = 2, cmsInfoCopyright = 3
#End If

'ICC Profiles describe transformations in specific color spaces.  A grayscale transform cannot be used on color data
' (and vice-versa).  We occasionally need to detect this information to prevent color space mismatches.
' (This is most commonly required at import time, if a color image file mistakenly has a grayscale ICC profile attached.)
'
'These values are defined on page 49 of the v2.8 LittleCMS API manual.
Public Enum LCMS_PROFILE_COLOR_SPACE
    cmsSigXYZ = 1482250784      'XYZ '
    cmsSigLab = 1281450528      'Lab '
    cmsSigLuv = 1282766368      'Luv '
    cmsSigYCbCr = 1497588338    'YCbr'
    cmsSigYxy = 1501067552      'Yxy '
    cmsSigRgb = 1380401696      'RGB '
    cmsSigGray = 1196573017     'GRAY '
    cmsSigHsv = 1213421088      'HSV '
    cmsSigHls = 1212961568      'HLS '
    cmsSigCmyk = 1129142603     'CMYK'
    cmsSigCmy = 1129142560      'CMY '
    cmsSigMCH1 = 1296255025     'MCH1'
    cmsSigMCH2 = 1296255026     'MCH2'
    cmsSigMCH3 = 1296255027     'MCH3'
    cmsSigMCH4 = 1296255028     'MCH4'
    cmsSigMCH5 = 1296255029     'MCH5'
    cmsSigMCH6 = 1296255030     'MCH6'
    cmsSigMCH7 = 1296255031     'MCH7'
    cmsSigMCH8 = 1296255032     'MCH8'
    cmsSigMCH9 = 1296255033     'MCH9'
    cmsSigMCHA = 1296255034     'MCHA'
    cmsSigMCHB = 1296255035     'MCHB'
    cmsSigMCHC = 1296255036     'MCHC'
    cmsSigMCHD = 1296255037     'MCHD'
    cmsSigMCHE = 1296255038     'MCHE'
    cmsSigMCHF = 1296255039     'MCHF'
    cmsSigNamed = 1852662636    'nmcl'
    cmsSig1color = 826494034    '1CLR'
    cmsSig2color = 843271250    '2CLR'
    cmsSig3color = 860048466    '3CLR'
    cmsSig4color = 876825682    '4CLR'
    cmsSig5color = 893602898    '5CLR'
    cmsSig6color = 910380114    '6CLR'
    cmsSig7color = 927157330    '7CLR'
    cmsSig8color = 943934546    '8CLR'
    cmsSig9color = 960711762    '9CLR'
    cmsSig10color = 1094929490  'ACLR'
    cmsSig11color = 1111706706  'BCLR'
    cmsSig12color = 1128483922  'CCLR'
    cmsSig13color = 1145261138  'DCLR'
    cmsSig14color = 1162038354  'ECLR'
    cmsSig15color = 1178815570  'FCLR'
    cmsSigLuvK = 1282766411     'LuvK'
End Enum

#If False Then
    Private Const cmsSigXYZ = 1482250784, cmsSigLab = 1281450528, cmsSigLuv = 1282766368, cmsSigYCbCr = 1497588338, cmsSigYxy = 1501067552, cmsSigRgb = 1380401696, cmsSigGray = 1196573017, cmsSigHsv = 1213421088, cmsSigHls = 1212961568, cmsSigCmyk = 1129142603, cmsSigCmy = 1129142560, cmsSigMCH1 = 1296255025, cmsSigMCH2 = 1296255026, cmsSigMCH3 = 1296255027, cmsSigMCH4 = 1296255028, cmsSigMCH5 = 1296255029, cmsSigMCH6 = 1296255030, cmsSigMCH7 = 1296255031, cmsSigMCH8 = 1296255032, cmsSigMCH9 = 1296255033, cmsSigMCHA = 1296255034, cmsSigMCHB = 1296255035, cmsSigMCHC = 1296255036, cmsSigMCHD = 1296255037, cmsSigMCHE = 1296255038, cmsSigMCHF = 1296255039, cmsSigNamed = 1852662636, cmsSig1color = 826494034, cmsSig2color = 843271250, cmsSig3color = 860048466, cmsSig4color = 876825682, cmsSig5color = 893602898, cmsSig6color = 910380114, cmsSig7color = 927157330, cmsSig8color = 943934546, cmsSig9color = 960711762, cmsSig10color = 1094929490, cmsSig11color = 1111706706
    Private Const cmsSig12color = 1128483922, cmsSig13color = 1145261138, cmsSig14color = 1162038354, cmsSig15color = 1178815570, cmsSigLuvK = 1282766411
#End If

'Unicode-aware LCMS functions require three-char ISO language and region names.  We use a dummy struct
' to simplify the process of enforcing a trailing null char, regardless of lang/region name length.
Private Type ThreeAsciiChars
    Chars(0 To 3) As Byte
End Type

'Return the current library version as a Long, e.g. "2.7" is returned as "2070"
Private Declare Function cmsGetEncodedCMMversion Lib "lcms2" () As Long

'Error logger registration; note that lcms must be custom-built to ensure this function signature is
' explicitly marked as stdcall; a default build assumes cdecl.
Private Declare Sub cmsSetLogErrorHandler Lib "lcms2" (ByVal ptrToCmsLogErrorHandlerFunction As Long)

'Profile create/release functions; white points declared as ByVal Longs can typically be set to NULL to use the default D50 value
Private Declare Function cmsCloseProfile Lib "lcms2" (ByVal srcProfile As Long) As Long
Private Declare Function cmsCreateBCHSWabstractProfile Lib "lcms2" (ByVal nLUTPoints As Long, ByVal newBrightness As Double, ByVal newContrast As Double, ByVal newHue As Double, ByVal newSaturation As Double, ByVal srcTemp As Long, ByVal dstTemp As Long) As Long
Private Declare Function cmsCreateGrayProfile Lib "lcms2" (ByVal ptrToWhitePointxyY As Long, ByVal sourceToneCurve As Long) As Long
Private Declare Function cmsCreateLab2Profile Lib "lcms2" (ByVal ptrToWhitePointxyY As Long) As Long
Private Declare Function cmsCreateLab4Profile Lib "lcms2" (ByVal ptrToWhitePointxyY As Long) As Long
Private Declare Function cmsCreate_sRGBProfile Lib "lcms2" () As Long
Private Declare Function cmsCreateRGBProfile Lib "lcms2" (ByVal ptrToWhitePointxyY As Long, ByVal ptrTo3xyYPrimaries As Long, ByVal ptrTo3ToneCurves As Long) As Long
'Private Declare Function cmsCreateXYZProfile Lib "lcms2" () As Long
Private Declare Function cmsOpenProfileFromMem Lib "lcms2" (ByVal ptrProfile As Long, ByVal profileSizeInBytes As Long) As Long
Private Declare Function cmsSaveProfileToMem Lib "lcms2" (ByVal srcProfile As Long, ByVal dstPtr As Long, ByRef sizeRequiredInBytes As Long) As Long

'Profile information functions
'Private Declare Function cmsGetEncodedICCversion Lib "lcms2" (ByVal hProfile As Long) As Long
Private Declare Function cmsGetHeaderRenderingIntent Lib "lcms2" (ByVal hProfile As Long) As LCMS_RENDERING_INTENT
Private Declare Function cmsGetProfileInfo Lib "lcms2" (ByVal hProfile As Long, ByVal srcInfo As LCMS_INFOTYPE, ByVal ptrToLanguageCode As Long, ByVal ptrToCountryCode As Long, ByVal ptrToWCharBuffer As Long, ByVal necessaryBufferSize As Long) As Long
Private Declare Function cmsGetProfileVersion Lib "lcms2" (ByVal hProfile As Long) As Double
Private Declare Function cmsGetPCS Lib "lcms2" (ByVal hProfile As Long) As LCMS_PROFILE_COLOR_SPACE
Private Declare Function cmsGetColorSpace Lib "lcms2" (ByVal hProfile As Long) As LCMS_PROFILE_COLOR_SPACE
Private Declare Sub cmsSetProfileVersion Lib "lcms2" (ByVal hProfile As Long, ByVal newVersion As Double)
 
'Tone curve creation/destruction.  (Not all are used right now.)
Private Declare Function cmsBuildParametricToneCurve Lib "lcms2" (ByVal contextID As Long, ByVal tcType As Long, ByVal ptrToFirstParam As Long) As Long
Private Declare Function cmsBuildGamma Lib "lcms2" (ByVal contextID As Long, ByVal gammaValue As Double) As Long
Private Declare Sub cmsFreeToneCurve Lib "lcms2" (ByVal srcToneCurve As Long)

'Transform functions
Private Declare Function cmsCreateTransform Lib "lcms2" (ByVal hInputProfile As Long, ByVal hInputFormat As LCMS_PIXEL_FORMAT, ByVal hOutputProfile As Long, ByVal hOutputFormat As LCMS_PIXEL_FORMAT, ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT, ByVal trnsFlags As LCMS_TRANSFORM_FLAGS) As Long
Private Declare Function cmsCreateMultiprofileTransform Lib "lcms2" (ByVal ptrToFirstProfile As Long, ByVal numOfProfiles As Long, ByVal hInputFormat As LCMS_PIXEL_FORMAT, ByVal hOutputFormat As LCMS_PIXEL_FORMAT, ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT, ByVal trnsFlags As LCMS_TRANSFORM_FLAGS) As Long
Private Declare Sub cmsDeleteTransform Lib "lcms2" (ByVal hTransform As Long)

'Color space conversions; any conversion that requires an XYZ WhitePoint can pass NULL
' for default D50 values
'Private Declare Sub cmsLab2XYZ Lib "lcms2" (ByVal ptrToWhitePointXYZ As Long, ByRef dstXYZ As LCMS_XYZ, ByRef srcLab As LCMS_Lab)
'Private Declare Sub cmsXYZ2Lab Lib "lcms2" (ByVal ptrToWhitePointXYZ As Long, ByRef dstLab As LCMS_Lab, ByRef srcXYZ As LCMS_XYZ)
'Private Declare Sub cmsXYZ2xyY Lib "lcms2" (ByRef dstxyY As LCMS_xyY, ByRef srcXYZ As LCMS_XYZ)
'Private Declare Sub cmsxyY2XYZ Lib "lcms2" (ByRef dstXYZ As LCMS_XYZ, ByRef srcxyY As LCMS_xyY)
Private Declare Function cmsWhitePointFromTemp Lib "lcms2" (ByRef dstWhitePointxyY As LCMS_xyY, ByVal srcTemperature As Double) As Long
'Private Declare Function cmsTempFromWhitePoint Lib "lcms2" (ByRef dstTemperature As Double, ByRef srcWhitePointxyY As LCMS_xyY) As Long

'Pointers to the constant XYZ/xyY declarations for D50
'Private Declare Function cmsD50_XYZ Lib "lcms2" () As Long
Private Declare Function cmsD50_xyY Lib "lcms2" () As Long

'Similar internal functions for D65 (which is used by a number of RGB spaces, e.g. Adobe and sRGB)
Private m_D65_XYZ() As Double, m_D65_xyY() As Double

'Actual transform application functions
Private Declare Sub cmsDoTransform Lib "lcms2" (ByVal hTransform As Long, ByVal ptrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal numOfPixelsToTransform As Long)

'In 2.8, a dedicated line/stride transform function was added to LittleCMS.  Here is what the documentation says:
' "This function translates bitmaps with complex organization. Each bitmap may contain several lines, and every line
'  may have padding. The distance from one line to the next one is BytesPerLine{In/Out}.  In planar formats, each line
'  may hold several planes, and each plane may have padding. Padding of lines and planes should be same across the whole
'  bitmap, i.e. all lines in a bitmap must be padded the same way. This function may be more efficient that repeated calls
'  to cmsDoTransform(), especially when customized plug-ins are being used."
'
'I do not currently make use of this function, but given the efficiency caveat above, it may be worth investigating in the future.
Private Declare Sub cmsDoTransformLineStride Lib "lcms2" (ByVal hTransform As Long, ByVal ptrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal numOfPixelsPerLine As Long, ByVal numOfLines As Long, ByVal bytesPerLineIn As Long, ByVal bytesPerLineOut As Long, ByVal bytesPerPlaneIn As Long, ByVal bytesPerPlaneOut As Long)

'A single LittleCMS handle is maintained for the life of a PD instance; see InitializeLCMS and ReleaseLCMS, below.
Private m_LCMSHandle As Long

'Initialize LittleCMS.  Do not call this until you have verified the LCMS plugin's existence
' (typically via the PluginManager module)
Public Function InitializeLCMS() As Boolean
    
    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim lcmsPath As String
    lcmsPath = PluginManager.GetPluginPath & "lcms2.dll"
    m_LCMSHandle = VBHacks.LoadLib(lcmsPath)
    InitializeLCMS = (m_LCMSHandle <> 0)
    
    If (Not InitializeLCMS) Then
        PDDebug.LogAction "WARNING!  LoadLibrary failed to load LittleCMS.  Last DLL error: " & Err.LastDllError
        PDDebug.LogAction "(FYI, the attempted path was: " & lcmsPath & ")"
    Else
        
        'NOTE: error callbacks are not reliable, so do not enable this in production code.  It can be
        ' useful, however, for tracking down esoteric errors in debug builds.
        If (PD_BUILD_QUALITY <> PD_PRODUCTION) Then cmsSetLogErrorHandler AddressOf LCMS_ErrorCallback
        
    End If
    
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
    VBHacks.FreeLib m_LCMSHandle
    m_LCMSHandle = 0
    PluginManager.SetPluginEnablement CCP_LittleCMS, False
End Sub

'After LittleCMS has been initialized, you can call this function to retrieve its current version.
' The version will always be formatted as "Major.Minor.0.0".
Public Function GetLCMSVersion() As String
    
    Dim versionAsLong As Long
    versionAsLong = cmsGetEncodedCMMversion()
    
    'The version is encoded as a 4-digit long, so e.g. 2.12.0 is "2120".
    ' This versioning mechanism has been valid since this 2.0 release (which was v. "2000")
    ' so I do not worry about it changing in the future.
    Dim versionAsString As String
    If (versionAsLong >= 1000) Then
        versionAsString = Left$(CStr(versionAsLong), 1) & "."
        versionAsString = versionAsString & Mid$(CStr(versionAsLong), 2, 2) & "."
        versionAsString = versionAsString & Right$(CStr(versionAsLong), 1) & ".0"
    Else
        versionAsString = "0.0.0"
    End If
    
    GetLCMSVersion = versionAsString
    
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
        LCMS_GetIlluminantxyY = (cmsWhitePointFromTemp(dstxyY, srcTemperature) <> 0)
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

Public Function LCMS_GetProfileInfoString(ByVal hInputProfile As Long, ByVal profileInfoType As LCMS_INFOTYPE, Optional ByVal languageCode As String = "en", Optional ByVal countryCode As String = "US") As String
    
    If (hInputProfile <> 0) Then
    
        'LCMS requires ISO country and region codes for locale-aware information.  For now, default to en-US.
        ' (In the future, it would be nice to change these to match the current program language.)
        Dim langCode As ThreeAsciiChars, countCode As ThreeAsciiChars
        
        Dim i As Long, lenInput As Long
        lenInput = Len(languageCode)
        If (lenInput > 3) Then lenInput = 3
        
        For i = 0 To lenInput - 1
            langCode.Chars(i) = Asc(Mid$(languageCode, i + 1, 1))
        Next i
       
        lenInput = Len(countryCode)
        If (lenInput > 3) Then lenInput = 3
       
        For i = 0 To lenInput - 1
            countCode.Chars(i) = Asc(Mid$(countryCode, i + 1, 1))
        Next i
        
        'Start by retrieving the length of the requested information
        Dim infoLength As Long
        infoLength = cmsGetProfileInfo(hInputProfile, profileInfoType, VarPtr(langCode.Chars(0)), VarPtr(countCode.Chars(0)), 0&, 0&)
        
        'If the length is non-zero, retrieve the full information string
        If (infoLength <> 0) Then
            LCMS_GetProfileInfoString = Space$(infoLength \ 2)
            cmsGetProfileInfo hInputProfile, profileInfoType, VarPtr(langCode.Chars(0)), VarPtr(countCode.Chars(0)), StrPtr(LCMS_GetProfileInfoString), infoLength
        End If
        
    End If
    
End Function

Public Function LCMS_CreateTwoProfileTransform(ByVal hInputProfile As Long, ByVal hOutputProfile As Long, Optional ByVal hInputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal hOutputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = cmsFLAGS_COPY_ALPHA) As Long
    trnsFlags = ValidateAlphaFlags(hInputFormat, trnsFlags)
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
    
    trnsFlags = ValidateAlphaFlags(pxFormat, trnsFlags)
    LCMS_CreateInPlaceTransformForDIB = cmsCreateTransform(hInputProfile, pxFormat, hOutputProfile, pxFormat, trnsRenderingIntent, trnsFlags)
    
End Function

'Validate COPY_ALPHA flag against source color space
Private Function ValidateAlphaFlags(ByVal inFormat As LCMS_PIXEL_FORMAT, ByVal inFlags As LCMS_TRANSFORM_FLAGS) As LCMS_TRANSFORM_FLAGS
    
    ValidateAlphaFlags = inFlags
    
    'In lcms 2.16, new behavior was implemented for the COPY_ALPHA flag.  Alpha *must* be present in the source format
    ' (destination doesn't matter) or this flag will cause a crash.
    '
    'Rather than add checks to *every* color-management call in PD (there are a ton!), perform a universal check here.
    If ((inFormat And FLAG_ALPHAPRESENT) = 0) And ((inFlags And cmsFLAGS_COPY_ALPHA) <> 0) Then
        ValidateAlphaFlags = ValidateAlphaFlags And (Not cmsFLAGS_COPY_ALPHA)
    End If
    
End Function

Public Function LCMS_DeleteTransform(ByRef hTransform As Long) As Boolean
    cmsDeleteTransform hTransform
    hTransform = 0
    LCMS_DeleteTransform = True
End Function

Public Function LCMS_GetProfileColorSpace(ByVal hProfile As Long) As LCMS_PROFILE_COLOR_SPACE
    LCMS_GetProfileColorSpace = cmsGetColorSpace(hProfile)
End Function

Public Function LCMS_GetProfileConnectionSpace(ByVal hProfile As Long) As LCMS_PROFILE_COLOR_SPACE
    LCMS_GetProfileConnectionSpace = cmsGetPCS(hProfile)
End Function

Public Function LCMS_GetProfileRenderingIntent(ByVal hProfile As Long) As LCMS_RENDERING_INTENT
    LCMS_GetProfileRenderingIntent = cmsGetHeaderRenderingIntent(hProfile)
End Function

Public Function LCMS_GetProfileVersion(ByVal hProfile As Long) As Double
    LCMS_GetProfileVersion = cmsGetProfileVersion(hProfile)
End Function

Public Function LCMS_LoadProfileFromMemory(ByVal ptrToProfile As Long, ByVal sizeOfProfileInBytes As Long) As Long
    LCMS_LoadProfileFromMemory = cmsOpenProfileFromMem(ptrToProfile, sizeOfProfileInBytes)
End Function

'Little CMS has its own "load from file" function, but it isn't Unicode-aware, so we just slam the file into a byte array
' and use the "load from memory" function instead.
Public Function LCMS_LoadProfileFromFile(ByRef profilePath As String) As Long
    
    LCMS_LoadProfileFromFile = 0

    Dim tmpProfileArray() As Byte
    If Files.FileExists(profilePath) Then
        If Files.FileLoadAsByteArray(profilePath, tmpProfileArray) Then LCMS_LoadProfileFromFile = cmsOpenProfileFromMem(VarPtr(tmpProfileArray(0)), UBound(tmpProfileArray) + 1)
    End If
    
End Function

'Create a custom RGB profile on-the-fly, using the specified white point and primaries.  No validation is
' performed on input values (except for the minimal validation applied automatically by LittleCMS), so please
' ensure that values are correct *before* calling this function.
Public Function LCMS_LoadCustomRGBProfile(ByVal ptrToWhitePointxyY As Long, ByVal ptrTo3xyYPrimaries As Long, Optional ByVal gammaCorrectFactor As Double = 1#) As Long

    'Use the supplied gamma value to generate LCMS-specific tone curves
    Dim rgbToneCurves() As Long
    ReDim rgbToneCurves(0 To 2) As Long
    rgbToneCurves(0) = LCMS_GetBasicToneCurve(gammaCorrectFactor)
    rgbToneCurves(1) = rgbToneCurves(0)
    rgbToneCurves(2) = rgbToneCurves(0)
    
    LCMS_LoadCustomRGBProfile = cmsCreateRGBProfile(ptrToWhitePointxyY, ptrTo3xyYPrimaries, VarPtr(rgbToneCurves(0)))
    
    'The intermediate tone curve *must* be freed now, as it's never directly exposed to the caller
    LCMS_FreeToneCurve rgbToneCurves(0)
    
End Function

'Create a custom RGB profile on-the-fly, using the specified white point and primaries.  No validation is
' performed on input values (except for the minimal validation applied automatically by LittleCMS), so please
' ensure that values are correct *before* calling this function.
'
'This "advanced" variant allows you to pass your own tone curves for each channel
Public Function LCMS_LoadCustomRGBProfile_Advanced(ByVal ptrToWhitePointxyY As Long, ByVal ptrTo3xyYPrimaries As Long, ByVal pSrcToneCurveR As Long, ByVal pSrcToneCurveG As Long, ByVal pSrcToneCurveB As Long, Optional ByVal freeCurvesForMe As Boolean = True) As Long

    'Use the supplied gamma value to generate LCMS-specific tone curves
    Dim rgbToneCurves() As Long
    ReDim rgbToneCurves(0 To 2) As Long
    rgbToneCurves(0) = pSrcToneCurveR
    rgbToneCurves(1) = pSrcToneCurveG
    rgbToneCurves(2) = pSrcToneCurveB
    
    LCMS_LoadCustomRGBProfile_Advanced = cmsCreateRGBProfile(ptrToWhitePointxyY, ptrTo3xyYPrimaries, VarPtr(rgbToneCurves(0)))
    
    'Free the source curve object if the user requests it (PD always frees a curve
    ' after it's been attached to a new profile)
    If freeCurvesForMe Then
        If (pSrcToneCurveB <> pSrcToneCurveG) And (pSrcToneCurveB <> pSrcToneCurveR) Then LCMS_FreeToneCurve pSrcToneCurveB
        If (pSrcToneCurveG <> pSrcToneCurveR) Then LCMS_FreeToneCurve pSrcToneCurveG
        LCMS_FreeToneCurve pSrcToneCurveR
    End If
    
End Function

Public Function LCMS_LoadStockGrayProfile(Optional ByVal useGamma As Double = 1#) As Long
    Dim tmpToneCurve As Long
    tmpToneCurve = LCMS_GetBasicToneCurve(useGamma)
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

Public Function LCMS_LoadStockSRGBProfile(Optional ByVal useIccV4 As Boolean = True) As Long
    LCMS_LoadStockSRGBProfile = cmsCreate_sRGBProfile()
    If (Not useIccV4) Then cmsSetProfileVersion LCMS_LoadStockSRGBProfile, 2.1
End Function

Public Function LCMS_LoadStockLabProfile(Optional ByVal useVersion4 As Boolean = True) As Long
    If useVersion4 Then
        LCMS_LoadStockLabProfile = cmsCreateLab4Profile(0&)
    Else
        LCMS_LoadStockLabProfile = cmsCreateLab2Profile(0&)
    End If
End Function

Public Function LCMS_SaveProfileToArray(ByVal hProfile As Long, ByRef dstArray() As Byte) As Boolean
    
    'Passing a null pointer will fill the "profile size" parameter with the required destination size
    Dim profSize As Long
    If (cmsSaveProfileToMem(hProfile, 0, profSize) <> 0) Then
        ReDim dstArray(0 To profSize - 1) As Byte
        LCMS_SaveProfileToArray = (cmsSaveProfileToMem(hProfile, VarPtr(dstArray(0)), profSize) <> 0)
    Else
        LCMS_SaveProfileToArray = False
    End If
    
End Function

Public Function LCMS_CreateAbstractBCHSProfile(Optional ByVal newBrightness As Double = 0#, Optional ByVal newContrast As Double = 1#, Optional ByVal newHue As Double = 0#, Optional ByVal newSaturation As Double = 0#, Optional ByVal srcTemp As Long = 0, Optional ByVal dstTemp As Long = 0) As Long
    LCMS_CreateAbstractBCHSProfile = cmsCreateBCHSWabstractProfile(16, newBrightness, newContrast, newHue, newSaturation, srcTemp, dstTemp)
End Function

Public Function LCMS_CloseProfileHandle(ByRef srcHandle As Long) As Boolean
    LCMS_CloseProfileHandle = (cmsCloseProfile(srcHandle) <> 0)
    If LCMS_CloseProfileHandle Then srcHandle = 0
End Function

Private Function LCMS_GetBasicToneCurve(Optional ByVal srcGamma As Double = 1#) As Long
    LCMS_GetBasicToneCurve = cmsBuildGamma(0&, srcGamma)
End Function

'paramValues *must* be an initialized and properly filled array sized to the relevant size for
' the passed curve type, or this function will crash.
Public Function LCMS_GetAdvancedToneCurve(ByVal curveType As Long, ByRef paramValues() As Double) As Long
    LCMS_GetAdvancedToneCurve = cmsBuildParametricToneCurve(0&, curveType, VarPtr(paramValues(0)))
End Function

Public Function LCMS_FreeToneCurve(ByRef hCurve As Long) As Boolean
    cmsFreeToneCurve hCurve
    hCurve = 0
    LCMS_FreeToneCurve = True
End Function

'Apply an already-created transform to a pdDIB object.
Public Function LCMS_ApplyTransformToDIB(ByRef srcDIB As pdDIB, ByVal hTransform As Long) As Boolean
    
    If (Not srcDIB Is Nothing) And (hTransform <> 0) Then
        
        '32-bpp DIBs can be applied in one fell swoop, since there are no scanline padding issues
        If (srcDIB.GetDIBColorDepth = 32) Then
            cmsDoTransform hTransform, srcDIB.GetDIBPointer, srcDIB.GetDIBPointer, srcDIB.GetDIBWidth * srcDIB.GetDIBHeight
                    
        '24-bpp DIBs may have scanline padding issues.  We must process them one line at a time.
        Else
            
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

'Apply an already-created transform to some arbitrary sub-region of a pdDIB object.
' NOTE!  This function performs no validation on the incoming rect.  You *must* ensure that its bounds fit entirely inside
' the pdDIB object or this function will crash and burn.
Public Function LCMS_ApplyTransformToDIB_RectF(ByRef srcDIB As pdDIB, ByVal hTransform As Long, ByRef dstRectF As RectF) As Boolean
    
    If ((Not srcDIB Is Nothing) And (hTransform <> 0)) Then
        
        'Start by determining a few values generic to this DIB
        Dim i As Long, iWidth As Long, iScanWidth As Long, iScanStart As Long
        iWidth = srcDIB.GetDIBWidth
        iScanStart = srcDIB.GetDIBPointer
        iScanWidth = srcDIB.GetDIBStride
        
        'Next, calculate unique offsets and strides based on the passed rect
        Dim pxSizeBytes As Long
        pxSizeBytes = srcDIB.GetDIBColorDepth \ 8
        
        Dim perLineOffset As Long, perLinePixels As Long
        perLineOffset = (Int(dstRectF.Left) * pxSizeBytes)
        perLinePixels = Int(dstRectF.Width + PDMath.Frac(dstRectF.Left) + 0.9999)
        If (Int(dstRectF.Left) + perLinePixels > srcDIB.GetDIBWidth - 1) Then perLinePixels = (srcDIB.GetDIBWidth - 1) - Int(dstRectF.Left)
        
        Dim startLine As Long, stopLine As Long
        startLine = Int(dstRectF.Top)
        stopLine = Int(dstRectF.Top + dstRectF.Height + 0.9999)
        If (stopLine > srcDIB.GetDIBHeight - 1) Then stopLine = srcDIB.GetDIBHeight - 1
        
        For i = startLine To stopLine
            cmsDoTransform hTransform, iScanStart + i * iScanWidth + perLineOffset, iScanStart + i * iScanWidth + perLineOffset, perLinePixels
        Next i
        
        'The "cmsDoTransform" function has no return, so we assume success if passed a valid DIB and transform
        LCMS_ApplyTransformToDIB_RectF = True
        
    End If
        
End Function

Public Sub LCMS_TransformArbitraryMemory(ByVal srcPointer As Long, ByVal dstPointer As Long, ByVal imgWidthInPixels As Long, ByVal hTransform As Long)
    cmsDoTransform hTransform, srcPointer, dstPointer, imgWidthInPixels
End Sub

Public Sub LCMS_TransformArbitraryMemoryEx(ByVal hTransform As Long, ByVal srcPointer As Long, ByVal dstPointer As Long, ByVal imgWidthInPixels As Long, ByVal numOfLines As Long, ByVal bytesPerLineIn As Long, ByVal bytesPerLineOut As Long, ByVal bytesPerPlaneIn As Long, ByVal bytesPerPlaneOut As Long)
    cmsDoTransformLineStride hTransform, srcPointer, dstPointer, imgWidthInPixels, numOfLines, bytesPerLineIn, bytesPerLineOut, bytesPerPlaneIn, bytesPerPlaneOut
End Sub

Public Sub LCMS_ErrorCallback(ByVal cmsContextID As Long, ByVal cmsUInt32ErrorCode As Long, ByVal ptrToDescription As Long)
    PDDebug.LogAction "LCMS returned error #" & CStr(cmsUInt32ErrorCode) & ": " & Strings.StringFromCharPtr(ptrToDescription, False), PDM_External_Lib
End Sub

'Given a target DIB and a valid pdICCProfile object, apply said profile to said DIB.
' (NOTE!  If the source image is 32-bpp, with premultiplied alpha, you need to unpremultiply alpha prior to
'         calling this function; otherwise, the end result will be invalid.)
Public Function ApplyICCProfileToPDDIB(ByRef targetDIB As pdDIB, ByRef srcIccProfile As pdICCProfile) As Boolean
    
    ApplyICCProfileToPDDIB = False
    
    If (targetDIB Is Nothing) Or (srcIccProfile Is Nothing) Then
        PDDebug.LogAction "WARNING!  LittleCMS.ApplyICCProfileToPDDIB was passed a null image or profile."
        Exit Function
    End If
    
    'Before doing anything else, make sure we actually have an ICC profile to apply!
    If (Not srcIccProfile.HasICCData) Then
        PDDebug.LogAction "ICC transform requested, but no data found.  Abandoning attempt."
        Exit Function
    End If
    
    PDDebug.LogAction "Using embedded ICC profile to convert image to sRGB space for editing..."
    
    'Start by creating two LCMS profile handles:
    ' 1) a source profile (the in-memory copy of the ICC profile associated with this DIB)
    ' 2) a destination profile (the current PhotoDemon working space)
    Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile
    Set srcProfile = New pdLCMSProfile
    Set dstProfile = New pdLCMSProfile
    
    If srcProfile.CreateFromPDICCObject(srcIccProfile) Then
        
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
                    PDDebug.LogAction "ICC profile transformation successful.  Image now lives in the current RGB working space."
                    targetDIB.SetColorManagementState cms_ProfileConverted
                    'TODO: assign the target DIB the hash of the color profile?  We currently leave this to
                    ' the caller, which is possibly a better idea...
                    ApplyICCProfileToPDDIB = True
                End If
                
                'Note that we could free the transform here, but it's unnecessary.  (The pdLCMSTransform class
                ' is self-freeing upon destruction.)
                
            Else
                PDDebug.LogAction "WARNING!  LittleCMS.ApplyICCProfileToPDDIB failed to create a valid transformation handle!"
            End If
        
        Else
            PDDebug.LogAction "WARNING!  LittleCMS.ApplyICCProfileToPDDib failed to create a valid destination profile handle."
        End If
    
    Else
        PDDebug.LogAction "WARNING!  LittleCMS.ApplyICCProfileToPDDib failed to create a valid source profile handle."
    End If
    
End Function
