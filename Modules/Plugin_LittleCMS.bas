Attribute VB_Name = "LittleCMS"
'***************************************************************************
'LittleCMS Interface
'Copyright 2016-2016 by Tanner Helland
'Created: 21/April/16
'Last updated: 21/April/16
'Last update: initial build
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

Public Type LCMS_cmsCIEXYZ
    x As Double
    y As Double
    z As Double
End Type

Public Type LCMS_cmsCIExyY
    x As Double
    y As Double
    YY As Double
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
    cmsFLAGS_8BITS_DEVICELINK = &H8&              ' Create 8 bits devicelinks
    cmsFLAGS_GUESSDEVICECLASS = &H20&             ' Guess device class (for transform2devicelink)
    cmsFLAGS_KEEP_SEQUENCE = &H80&                ' Keep profile sequence for devicelink creation
    ' Specific to a particular optimizations
    cmsFLAGS_FORCE_CLUT = &H2&                     ' Force CLUT optimization
    cmsFLAGS_CLUT_POST_LINEARIZATION = &H1&        ' create postlinearization tables if possible
    cmsFLAGS_CLUT_PRE_LINEARIZATION = &H10&        ' create prelinearization tables if possible
    ' CRD special
    cmsFLAGS_NODEFAULTRESOURCEDEF = &H1000000
End Enum

'Return the current library version as a Long, e.g. "2.7" is returned as "2070"
Private Declare Function cmsGetEncodedCMMversion Lib "lcms2.dll" () As Long

'Error logger registration
Private Declare Sub cmsSetLogErrorHandler Lib "lcms2.dll" (ByVal ptrToCmsLogErrorHandlerFunction As Long)

'Profile create/release functions
Private Declare Function cmsCloseProfile Lib "lcms2.dll" (ByVal srcProfile As Long) As Long
Private Declare Function cmsCreateGrayProfile Lib "lcms2.dll" (ByRef whitePoint As LCMS_cmsCIExyY, ByVal sourceToneCurve As Long) As Long
Private Declare Function cmsCreate_sRGBProfile Lib "lcms2.dll" () As Long
Private Declare Function cmsOpenProfileFromMem Lib "lcms2.dll" (ByVal ptrProfile As Long, ByVal profileSizeInBytes As Long) As Long

'Profile information functions
Private Declare Function cmsGetHeaderRenderingIntent Lib "lcms2.dll" (ByVal hProfile As Long) As LCMS_RENDERING_INTENT

'Tone curve creation/destruction
Private Declare Function cmsBuildParametricToneCurve Lib "lcms2.dll" (ByVal ContextID As Long, ByVal tcType As Long, ByVal ptrToFirstParam As Long) As Long
Private Declare Function cmsBuildGamma Lib "lcms2.dll" (ByVal ContextID As Long, ByVal cmsFloat64Number As Double) As Long
Private Declare Sub cmsFreeToneCurve Lib "lcms2.dll" (ByVal srcToneCurve As Long)

'Transform functions
Private Declare Function cmsCreateTransform Lib "lcms2.dll" (ByVal hInputProfile As Long, ByVal hInputFormat As LCMS_PIXEL_FORMAT, ByVal hOutputProfile As Long, ByVal hOutputFormat As LCMS_PIXEL_FORMAT, ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT, ByVal trnsFlags As LCMS_TRANSFORM_FLAGS) As Long
Private Declare Sub cmsDeleteTransform Lib "lcms2.dll" (ByVal hTransform As Long)

'Actual transform application functions
Private Declare Sub cmsDoTransform Lib "lcms2.dll" (ByVal hTransform As Long, ByVal ptrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal numOfPixelsToTransform As Long)

'A single LittleCMS handle is maintained for the life of a PD instance; see InitializeLCMS and ReleaseLCMS, below.
Private m_LCMSHandle As Long

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
        'If InitializeLCMS Then
        '    cmsSetLogErrorHandler AddressOf cmsErrorHandler
        'End If
        
        If (Not InitializeLCMS) Then
            pdDebug.LogAction "WARNING!  LoadLibrary failed to load LittleCMS.  Last DLL error: " & Err.LastDllError
            pdDebug.LogAction "(FYI, the attempted path was: " & lcmsPath & ")"
        End If
    #End If
    
End Function

'When PD closes, make sure to release our library handle
Public Sub ReleaseLCMS()
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

Public Function LCMS_CreateTwoProfileTransform(ByVal hInputProfile As Long, ByVal hOutputProfile As Long, Optional ByVal hInputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal hOutputFormat As LCMS_PIXEL_FORMAT = TYPE_BGRA_8, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = 0) As Long
    LCMS_CreateTwoProfileTransform = cmsCreateTransform(hInputProfile, hInputFormat, hOutputProfile, hOutputFormat, trnsRenderingIntent, trnsFlags)
End Function

Public Function LCMS_CreateInPlaceTransformForDIB(ByVal hInputProfile As Long, ByVal hOutputProfile As Long, ByRef srcDIB As pdDIB, Optional ByVal trnsRenderingIntent As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL, Optional ByVal trnsFlags As LCMS_TRANSFORM_FLAGS = 0) As Long
    
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

Public Function LCMS_LoadStockGrayProfile() As Long
    Dim whitePoint As LCMS_cmsCIExyY
    With whitePoint
        .x = 0.3127
        .y = 0.329
        .YY = 1#
    End With
    
    Dim tmpToneCurve As Long
    tmpToneCurve = LCMS_GetBasicToneCurve(1#)
    LCMS_LoadStockGrayProfile = cmsCreateGrayProfile(whitePoint, tmpToneCurve)
    LittleCMS.LCMS_FreeToneCurve tmpToneCurve
End Function

Public Function LCMS_LoadStockSRGBProfile() As Long
    LCMS_LoadStockSRGBProfile = cmsCreate_sRGBProfile()
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

Public Function LCMS_ApplyTransformToDIB(ByRef srcDIB As pdDIB, ByVal hTransform As Long) As Boolean
    
    If (Not (srcDIB Is Nothing)) And (hTransform <> 0) Then
        
        '32-bpp DIBs can be applied in one fell swoop, since there are no scanline padding issues
        If (srcDIB.GetDIBColorDepth = 32) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Applying ICC transform to 32-bpp DIB..."
            #End If
            
            cmsDoTransform hTransform, srcDIB.GetActualDIBBits, srcDIB.GetActualDIBBits, srcDIB.GetDIBWidth * srcDIB.GetDIBHeight
        
        '24-bpp DIBs may have scanline padding issues.  We must process them one line at a time.
        Else
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Applying ICC transform to 24-bpp DIB..."
            #End If
            
            Dim i As Long, iWidth As Long, iScanWidth As Long, iScanStart As Long
            iWidth = srcDIB.GetDIBWidth
            iScanStart = srcDIB.GetActualDIBBits
            iScanWidth = srcDIB.GetDIBArrayWidth
            
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
        pdDebug.LogAction "WARNING!  LittleCMS returned the following error: #" & cmsError
    #End If
End Function
