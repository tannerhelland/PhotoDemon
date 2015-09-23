Attribute VB_Name = "Color_Management"
'***************************************************************************
'PhotoDemon ICC (International Color Consortium) Profile Support Module
'Copyright 2013-2015 by Tanner Helland
'Created: 05/November/13
'Last updated: 05/September/14
'Last update: tie the multiprofile transform quality to the new Color Management Performance preference
'
'ICC profiles can be embedded in certain image file formats.  These profiles can be used to convert an image into
' a precisely defined reference space, while taking into account any pecularities of the device that captured the
' image (typically a camera).  From that reference space, we can then convert the image into any other
' device-specific color space (typically a monitor or printer).
'
'ICC profile handling is broken into three parts: extracting the profile from an image, using the extracted profile to
' convert an image into a reference space (currently sRGB only), and then activating color management for any
' user-facing DCs using the color profiles specified by the user.  The extraction step is currently handled via
' FreeImage or GDI+, while the application step is handled by Windows.  In the future I may look at adding ExifTool as
' a possible mechanism for extracting the profile, as it provides better support for esoteric formats than FreeImage.
'
'This class does not perform the extraction of ICC Profile data from images.  That is handled by the pdICCProfile
' class, which operates on a per-image basis.  This module simply supplies a number of generic ICC-related functions.
'
'This module would not be possible without this excellent test code from pro VB coder LaVolpe:
' http://www.vbforums.com/showthread.php?666143-RESOLVED-ICC-%28Color-Profiles%29
' Note that LaVolpe's code contains a number of errors, so if you're looking to build your own ICC implementation,
' I suggest basing it off my work instead.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'A handle (HMONITOR, specifically) to the main form's current monitor.  This value is updated by firing the
' checkParentMonitor() function, below.
Private currentMonitor As Long

'When the main form's monitor changes, this string will automatically be updated with the corresponding ICC
' profile path of that monitor (if the user has selected a custom one)
Private currentColorProfile As String

'ICC Profile header; this stores basic information about a given profile, and is use to interact with various
' ICC-related API functions.
Private Type ICC_PROFILE
    dwType As Long
    pProfileData As Long
    cbDataSize As Long
End Type

'There are two possible dwType values for an ICC_PROFILE variable; we use PROFILE_MEMBUFFER
Private Const PROFILE_FILENAME As Long = 1&
Private Const PROFILE_MEMBUFFER As Long = 2&

'We only need to read ICC data, not write it
Private Const PROFILE_READ As Long = 1&

'Other functions are welcome to share the profile data
Private Const FILE_SHARE_READ As Long = 1&

'We want the function to fail if the profile cannot be opened; do not simply create a blank profile in its place
Private Const OPEN_EXISTING As Long = 3&

'Windows only provides two standard color profiles: sRGB, and the current system default.  These are declared as
' public so that external functions can request either of them.
Public Const LCS_CALIBRATED_RGB As Long = &H0       'This constant is technically unsupported, and should not be used.  I'm including it here for testing purposes only.
Public Const LCS_sRGB As Long = &H73524742
Public Const LCS_WINDOWS_COLOR_SPACE As Long = &H57696E20

'Profile transformation is not lossless; for example, it is rarely possible to perfectly preserve hue, saturation,
' and luminance - some components must be sacrificed in order to ideally render others.  By default, PhotoDemon uses
' the standard intent for displays, which is IntentPerceptual (basically, stretch the image's luminance so that its
' full gamut is viewable). These intents are declared as public so that external functions can request whichever
' render intent they desire for a given application.
Public Enum RenderingIntents
    INTENT_PERCEPTUAL = 0&
    INTENT_RELATIVECOLORIMETRIC = 1&
    INTENT_SATURATION = 2&
    INTENT_ABSOLUTECOLORIMETRIC = 3&
End Enum

#If False Then
    Const INTENT_PERCEPTUAL As Long = 0&, INTENT_RELATIVECOLORIMETRIC As Long = 1&, INTENT_SATURATION As Long = 2&, INTENT_ABSOLUTECOLORIMETRIC As Long = 3&
#End If

'Windows provides different qualities for profile transformations (proof, normal, best).  As we only use two-component
' transforms, performance isn't a crucial issue, so we use BEST by default.
Private Enum CMM_TRANSFORM_QUALITY
    PROOF_MODE = 1&
    NORMAL_MODE = 2&
    BEST_MODE = 3&
End Enum

#If False Then
    Const PROOF_MODE = 1&, NORMAL_MODE = 2&, BEST_MODE = 3&
#End If

'Because we only do ICC-to-ICC transforms, Windows can be instructed to use a 3rd-party CMS instead of its own
' internal one.  We don't care if it does this, and we tell it as much.
Private Const INDEX_DONT_CARE As Long = 0&

'When it comes time to actually apply the transformation to the image data, the transform needs to know the image's
' color depth.  PD only operates in 24 and 32bpp mode, so we only need two constants here.
Public Const BM_RGBTRIPLETS As Long = &H2
Public Const BM_BGRTRIPLETS As Long = &H4
Public Const BM_xBGRQUADS As Long = &H10
Public Const BM_xRGBQUADS As Long = &H8
Public Const BM_CMYKQUADS As Long = &H20
Public Const BM_KYMCQUADS As Long = &H305

'Various ICC-related APIs are needed to open profiles and transform data between them
Private Declare Function OpenColorProfile Lib "mscms" Alias "OpenColorProfileA" (ByRef pProfile As Any, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal dwCreationMode As Long) As Long
Private Declare Function CloseColorProfile Lib "mscms" (ByVal hProfile As Long) As Long
Private Declare Function IsColorProfileValid Lib "mscms" (ByVal hProfile As Long, ByRef pBool As Long) As Long
Private Declare Function GetStandardColorSpaceProfile Lib "mscms" Alias "GetStandardColorSpaceProfileA" (ByVal pcStr As String, ByVal dwProfileID As Long, ByVal pProfileName As Long, ByRef pdwSize As Long) As Long
Private Declare Function CreateMultiProfileTransform Lib "mscms" (ByRef pProfile As Any, ByVal nProfiles As Long, ByRef pIntents As Long, ByVal nIntents As Long, ByVal dwFlags As Long, ByVal indexPreferredCMM As Long) As Long
Private Declare Function DeleteColorTransform Lib "mscms" (ByVal hTransform As Long) As Long
Private Declare Function TranslateBitmapBits Lib "mscms" (ByVal hTransform As Long, ByVal srcBitsPointer As Long, ByVal pBmInput As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal dwInputStride As Long, ByVal dstBitsPointer As Long, ByVal pBmOutput As Long, ByVal dwOutputStride As Long, ByRef pfnCallback As Long, ByVal ulCallbackData As Long) As Long
Private Declare Function GetColorDirectory Lib "mscms" Alias "GetColorDirectoryA" (ByVal pMachineName As Long, ByVal pBuffer As Long, ByRef pdwSize As Long) As Long
Private Declare Function GetColorProfileHeader Lib "mscms" (ByVal pProfileHandle As Long, ByVal pHeaderBufferPointer As Long) As Long

'Windows handles color management on a per-DC basis.  Use SetICMMode and these constants to activate/deactivate or query a DC.
Private Declare Function SetICMMode Lib "gdi32" (ByVal targetDC As Long, ByVal iEnableICM As ICM_Mode) As Long

Private Enum ICM_Mode
    ICM_OFF = 1
    ICM_ON = 2
    ICM_QUERY = 3
    ICM_DONE_OUTSIDEDC = 4
End Enum

#If False Then
    Const ICM_OFF = 1, ICM_ON = 2, ICM_QUERY = 3, ICM_DONE_OUTSIDEDC = 4
#End If

'Retrieves the filename of the color management file associated with a given DC
Private Declare Function GetICMProfile Lib "gdi32" Alias "GetICMProfileA" (ByVal hDC As Long, ByRef lpcbName As Long, ByRef BufferPtr As Long) As Long
Private Declare Function SetICMProfile Lib "gdi32" Alias "SetICMProfileA" (ByVal hDC As Long, ByVal lpFileName As String) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

'When PD is first loaded, the system's current color management file will be cached in this variable
Private currentSystemColorProfile As String
Public g_IsSystemColorProfileSRGB As Boolean

Private Const MAX_PATH As Long = 260

'Shorthand way to activate color management for anything with a DC
Public Sub TurnOnDefaultColorManagement(ByVal targetDC As Long, ByVal targetHwnd As Long)
    
    'Perform a quick check to see if we the target DC is requesting sRGB management.  If it is, we can skip
    ' color management entirely, because PD stores all RGB data in sRGB anyway.
    If Not (g_UseSystemColorProfile And g_IsSystemColorProfileSRGB) Then
        AssignDefaultColorProfileToObject targetHwnd, targetDC
        TurnOnColorManagementForDC targetDC
    End If
    
End Sub

'Retrieve the current system color profile directory
Public Function GetSystemColorFolder() As String

    'Prepare a blank string to receive the profile path
    Dim filenameLength As Long
    filenameLength = MAX_PATH
    
    Dim tmpPathString As String
    tmpPathString = ""
    
    Dim tmpPathBuffer() As Byte
    ReDim tmpPathBuffer(0 To filenameLength - 1) As Byte
    
    'Use the GetColorDirectory function to request the location of the system color folder
    If GetColorDirectory(0&, ByVal VarPtr(tmpPathBuffer(0)), filenameLength) = 0 Then
        GetSystemColorFolder = ""
    Else
    
        'Convert the returned byte array into a string
        tmpPathString = StrConv(tmpPathBuffer, vbUnicode)
        tmpPathString = TrimNull(tmpPathString)
                
        GetSystemColorFolder = tmpPathString
        
    End If

End Function

'Assign the default color profile (whether the system profile or the user profile) to any arbitrary object.  Note that the object
' MUST have an hWnd and an hDC property for this to work.
Public Sub AssignDefaultColorProfileToObject(ByVal objectHWnd As Long, ByVal objectHDC As Long)
    
    'If the current user setting is "use system color profile", our job is easy.
    If g_UseSystemColorProfile Then
        SetICMProfile objectHDC, currentSystemColorProfile
    Else
        
        'Use the form's containing monitor to retrieve a matching profile from the preferences file
        If Len(currentColorProfile) <> 0 Then
            SetICMProfile objectHDC, currentColorProfile
        Else
            SetICMProfile objectHDC, currentSystemColorProfile
        End If
        
    End If
    
    'If you would like to test this function on a standalone ICC profile (generally something bizarre, to help you know
    ' that the function is working), use something similar to the code below.
    'Dim TEST_ICM As String
    'TEST_ICM = "C:\PhotoDemon v4\PhotoDemon\no_sync\Images from testers\jpegs\ICC\WhackedRGB.icc"
    'SetICMProfile targetDC, TEST_ICM
    
End Sub

'When PD is first loaded, this function will be called, which caches the current color management file in use by the system
Public Sub CacheCurrentSystemColorProfile()
    
    currentSystemColorProfile = GetDefaultICCProfile()
    
    'As part of this step, we will also temporarily load the default system ICC profile, and check to see if it's sRGB.
    ' If it is, we can skip color management entirely, as all images are processed in sRGB.
    
    'Obtain a handle to the default system profile
    Dim sysProfileHandle As Long
    sysProfileHandle = LoadICCProfileFromFile(currentSystemColorProfile)
    
    If sysProfileHandle <> 0 Then
    
        'Obtain a handle to a stock sRGB profile.
        Dim srgbProfileHandle As Long
        srgbProfileHandle = LoadStandardICCProfile(LCS_sRGB)
        
        'Compare the two profiles
        If AreColorProfilesEqual(sysProfileHandle, srgbProfileHandle) Then
            g_IsSystemColorProfileSRGB = True
        Else
            g_IsSystemColorProfileSRGB = False
        End If
        
        'Release our profile handles
        ReleaseICCProfile sysProfileHandle
        ReleaseICCProfile srgbProfileHandle
        
    Else
        
        Debug.Print "System ICC profile couldn't be loaded.  Default color management is disabled for this session."
        g_IsSystemColorProfileSRGB = True
        
    End If
    
End Sub

'Returns the path to the default color mangement profile file (ICC or WCS) currently in use by the system.
Public Function GetDefaultICCProfile() As String

    'Prepare a blank string to receive the profile path
    Dim filenameLength As Long
    filenameLength = MAX_PATH
    
    Dim tmpPathString As String
    tmpPathString = ""
    
    Dim tmpPathBuffer() As Byte
    ReDim tmpPathBuffer(0 To filenameLength - 1) As Byte
    
    'Using the desktop DC as our reference, request the filename of the currently in-use ICM profile (which should be the system default)
    If GetICMProfile(GetDC(0), filenameLength, ByVal VarPtr(tmpPathBuffer(0))) = 0 Then
        GetDefaultICCProfile = ""
    Else
    
        'Convert the returned byte array into a string
        Dim i As Long
        For i = 0 To filenameLength - 1
            tmpPathString = tmpPathString & Chr(tmpPathBuffer(i))
        Next i
                
        GetDefaultICCProfile = tmpPathString
        
    End If
    
End Function

'Turn on color management for a specified device context
Public Sub TurnOnColorManagementForDC(ByVal dstDC As Long)
    SetICMMode dstDC, ICM_ON
End Sub

'Turn off color management for a specified device context
Public Sub TurnOffColorManagementForDC(ByVal dstDC As Long)
    SetICMMode dstDC, ICM_OFF
End Sub

'Given a valid iccProfileArray (such as one stored in a pdICCProfile class), convert it to an internal Windows color profile
' handle, validate it, and return the result.  Returns a non-zero handle if successful.
Public Function LoadICCProfileFromMemory(ByVal profileArrayPointer As Long, ByVal profileArraySize As Long) As Long

    'Start by preparing an ICC_PROFILE header to use with the color management APIs
    Dim srcProfileHeader As ICC_PROFILE
    srcProfileHeader.dwType = PROFILE_MEMBUFFER
    srcProfileHeader.pProfileData = profileArrayPointer
    srcProfileHeader.cbDataSize = profileArraySize
    
    'Use that header to open a reference to an internal Windows color profile (which is required by all ICC-related API)
    LoadICCProfileFromMemory = OpenColorProfile(srcProfileHeader, PROFILE_READ, FILE_SHARE_READ, OPEN_EXISTING)
    
    If LoadICCProfileFromMemory <> 0 Then
    
        'Validate the profile's XML as well; it is possible for a profile to be ill-formed, which means we cannot use it.
        Dim tmpCheck As Long
        If IsColorProfileValid(LoadICCProfileFromMemory, tmpCheck) = 0 Then
            Debug.Print "Color profile loaded succesfully, but XML failed to validate."
            CloseColorProfile LoadICCProfileFromMemory
            LoadICCProfileFromMemory = 0
        End If
        
    Else
        Debug.Print "ICC profile failed to load (OpenColorProfile failed with error #" & Err.LastDllError & ")."
    End If

End Function

'Given a valid ICC profile path, convert it to an internal Windows color profile handle, validate it,
' and return the result.  Returns a non-zero handle if successful.
Public Function LoadICCProfileFromFile(ByVal profilePath As String) As Long

    Dim cFile As pdFSO
    Set cFile = New pdFSO

    'Start by loading the specified path into a byte array
    Dim tmpProfileArray() As Byte
        
    If cFile.FileExist(profilePath) Then
        
        If Not cFile.LoadFileAsByteArray(profilePath, tmpProfileArray) Then
            LoadICCProfileFromFile = 0
            Exit Function
        End If
        
    Else
        LoadICCProfileFromFile = 0
        Exit Function
    End If

    'Next, prepare an ICC_PROFILE header to use with the color management APIs
    Dim srcProfileHeader As ICC_PROFILE
    srcProfileHeader.dwType = PROFILE_MEMBUFFER
    srcProfileHeader.pProfileData = VarPtr(tmpProfileArray(0))
    srcProfileHeader.cbDataSize = UBound(tmpProfileArray) + 1
    
    'Use that header to open a reference to an internal Windows color profile (which is required by all ICC-related API)
    LoadICCProfileFromFile = OpenColorProfile(srcProfileHeader, PROFILE_READ, FILE_SHARE_READ, OPEN_EXISTING)
    
    If LoadICCProfileFromFile <> 0 Then
    
        'Validate the profile's XML as well; it is possible for a profile to be ill-formed, which means we cannot use it.
        Dim tmpCheck As Long
        If IsColorProfileValid(LoadICCProfileFromFile, tmpCheck) = 0 Then
            Debug.Print "Color profile loaded succesfully, but XML failed to validate."
            CloseColorProfile LoadICCProfileFromFile
            LoadICCProfileFromFile = 0
        End If
        
    Else
        Debug.Print "ICC profile failed to load (OpenColorProfile failed with error #" & Err.LastDllError & ")."
    End If

End Function

'Request a standard ICC profile from the OS.  Windows only provides two standard color profiles: sRGB (LCS_sRGB), and whatever
' the system default currently is (LCS_WINDOWS_COLOR_SPACE).  While probably not necessary, this function also validates the
' requested profile, just to be safe.
Public Function LoadStandardICCProfile(ByVal profileID As Long) As Long

    'Start by preparing a header for the destination ICC profile
    Dim dstProfileHeader As ICC_PROFILE
    dstProfileHeader.dwType = PROFILE_FILENAME
    
    'We do not know the size of the requested profile in advance, so we must use a specialized call to
    ' GetStandardColorSpaceProfile, which will fill the last parameter with the size of the profile.
    GetStandardColorSpaceProfile vbNullString, profileID, 0&, dstProfileHeader.cbDataSize
        
    '.cbDataSize now contains the size of the required sRGB profile.  Prepare a dummy array to hold the received data.
    Dim dstICCData() As Byte
    ReDim dstICCData(0 To dstProfileHeader.cbDataSize - 1) As Byte
    dstProfileHeader.pProfileData = VarPtr(dstICCData(0))
    
    'Now that we have an array to contain the profile, we use GetStandardColorSpaceProfile to fill it
    GetStandardColorSpaceProfile vbNullString, profileID, dstProfileHeader.pProfileData, dstProfileHeader.cbDataSize
        
    'With a fully populated header, it is finally time to open an internal Windows version of the data!
    LoadStandardICCProfile = OpenColorProfile(dstProfileHeader, PROFILE_READ, FILE_SHARE_READ, OPEN_EXISTING)
    
    'It is highly unlikely (maybe even impossible?) for the system to return an invalid standard profile, but just to be
    ' safe, validate the XML.
    If LoadStandardICCProfile <> 0 Then
    
        'Validate the profile's XML as well; it is possible for a profile to be ill-formed, which means we cannot use it.
        Dim tmpCheck As Long
        If IsColorProfileValid(LoadStandardICCProfile, tmpCheck) = 0 Then
            Debug.Print "Standard color profile loaded succesfully, but XML failed to validate."
            CloseColorProfile LoadStandardICCProfile
            LoadStandardICCProfile = 0
        End If
        
    Else
        Debug.Print "Standard ICC profile failed to load (GetStandardColorSpaceProfile failed with error #" & Err.LastDllError & ")."
    End If

End Function

'This function is just a thin wrapper to CloseColorProfile; however, using it allows us to keep various color-management
' DLLs nicely encapsulated within this module.
Public Sub ReleaseICCProfile(ByVal profileHandle As Long)
    CloseColorProfile profileHandle
End Sub

'Given a source profile, destination profile, and rendering intent, return a compatible transformation handle.
Public Function RequestProfileTransform(ByVal srcProfile As Long, ByVal dstProfile As Long, ByVal preferredIntent As RenderingIntents, Optional ByVal useEmbeddedIntent As Long = -1) As Long

    'Next we need to prepare two matrices to supply to CreateMultiProfileTransform: one for ICC profiles themselves,
    ' and one for desired render intents.
    Dim profileMatrix(0 To 1) As Long, intentMatrix(0 To 1) As Long
    
    'The first row in the array contains the two profile pointers we've already acquired, in src/dest order
    profileMatrix(0) = srcProfile
    profileMatrix(1) = dstProfile
    
    'The second column in the array contains the render intents for the transformation.  Note that an option is available
    ' to use a preferred intent in an ICC profile, if one exists.
    
    'DISCLAIMER! Until this setting can be handled by preference, I now default to perceptual render intent.  This provides
    ' better results on most images, and is standard for PostScript workflows.  See http://fieryforums.efi.com/showthread.php/835-Rendering-Intent-Control-for-Embedded-Profiles.
    ' or https://developer.mozilla.org/en-US/docs/ICC_color_correction_in_Firefox, for example)
    useEmbeddedIntent = -1
    
    If useEmbeddedIntent > -1 Then
        intentMatrix(0) = useEmbeddedIntent
    
    'If the user does not want us to use the embedded intent in the source file, simply mimic the preferred destination intent.
    Else
        intentMatrix(0) = preferredIntent
    End If
    
    'The destination
    intentMatrix(1) = preferredIntent
    
    'We can now use our profile matrix to generate a transformation object, which we will use on the DIB itself.
    ' Note: the quality of the transform will affect the speed of the resulting transformation.  Windows supports 3 quality levels
    '       on the range [1, 3].  We map our internal g_ColorPerformance preference on the range [0, 2] to that range, and use it
    '       to transparently adjust the quality of the transform.
    RequestProfileTransform = CreateMultiProfileTransform(ByVal VarPtr(profileMatrix(0)), 2&, ByVal VarPtr(intentMatrix(0)), 2&, (2 - g_ColorPerformance) + 1, INDEX_DONT_CARE)
    
    If RequestProfileTransform = 0 Then
        Debug.Print "Requested color transformation could not be generated (Error #" & Err.LastDllError & ")."
    End If
    
End Function

'This function is just a thin wrapper to DeleteColorTransform; however, using it allows us to keep various color-management
' DLLs nicely encapsulated within this module.
Public Sub ReleaseColorTransform(ByVal transformHandle As Long)
    DeleteColorTransform transformHandle
End Sub

'Given a color transformation and a DIB, apply one to the other!  Returns TRUE if successful.
Public Function ApplyColorTransformToDIB(ByVal srcTransform As Long, ByRef dstDIB As pdDIB) As Boolean

    Dim transformCheck As Long
    
    With dstDIB
                
        'NOTE: note that I use BM_RGBTRIPLETS below, despite pdDIB DIBs most definitely being in BGR order.  This is an
        '       undocumented bug with Windows' color management engine!
        Dim bitDepthIdentifier As Long
        If .getDIBColorDepth = 24 Then bitDepthIdentifier = BM_RGBTRIPLETS Else bitDepthIdentifier = BM_xRGBQUADS
                
        'TranslateBitmapBits handles the actual transformation for us.
        transformCheck = TranslateBitmapBits(srcTransform, .getActualDIBBits, bitDepthIdentifier, .getDIBWidth, .getDIBHeight, .getDIBArrayWidth, .getActualDIBBits, bitDepthIdentifier, .getDIBArrayWidth, ByVal 0&, 0&)
        
    End With
    
    If transformCheck = 0 Then
        ApplyColorTransformToDIB = False
        
        'Error #2021 is ERROR_COLORSPACE_MISMATCH: "The specified transform does not match the bitmap's color space."
        ' This is a known error when the source image was in CMYK format, because FreeImage (or GDI+) will have
        ' automatically converted the image to RGB at load-time.  Because the ICC profile is CMYK-specific, Windows will
        ' not be able to apply it to the image, as it is no longer in CMYK format!
        If CLng(Err.LastDllError) = 2021 Then
            Debug.Print "Note: sRGB conversion already occurred."
        Else
            Debug.Print "ICC profile could not be applied.  Image remains in original profile. (Error #" & Err.LastDllError & ")."
        End If
        
    Else
        ApplyColorTransformToDIB = True
    End If

End Function

'Given a color transformation and two DIBs, fill one DIB with a transformed copy of the other!  Returns TRUE if successful.
Public Function ApplyColorTransformToTwoDIBs(ByVal srcTransform As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal srcFormat As Long, ByVal dstFormat As Long) As Boolean

    Dim transformCheck As Long
    
    'TranslateBitmapBits handles the actual transformation for us.
    transformCheck = TranslateBitmapBits(srcTransform, srcDIB.getActualDIBBits, srcFormat, srcDIB.getDIBWidth, srcDIB.getDIBHeight, srcDIB.getDIBArrayWidth, dstDIB.getActualDIBBits, dstFormat, dstDIB.getDIBArrayWidth, ByVal 0&, 0&)
    
    If transformCheck = 0 Then
        ApplyColorTransformToTwoDIBs = False
        
        'Error #2021 is ERROR_COLORSPACE_MISMATCH: "The specified transform does not match the bitmap's color space."
        ' This is a known error when the source image was in CMYK format, because FreeImage (or GDI+) will have
        ' automatically converted the image to RGB at load-time.  Because the ICC profile is CMYK-specific, Windows will
        ' not be able to apply it to the image, as it is no longer in CMYK format!
        If CLng(Err.LastDllError) = 2021 Then
            Debug.Print "ICC profile could not be applied, because requested color spaces did not match supplied profile spaces."
        Else
            Debug.Print "ICC profile could not be applied.  Image remains in original profile. (Error #" & Err.LastDllError & ")."
        End If
        
    Else
        ApplyColorTransformToTwoDIBs = True
    End If

End Function

'Apply a CMYK transform between a 32bpp CMYK DIB and a 24bpp sRGB DIB.
Public Function ApplyCMYKTransform(ByVal iccProfilePointer As Long, ByVal iccProfileSize As Long, ByRef srcCMYKDIB As pdDIB, ByRef dstRGBDIB As pdDIB, Optional ByVal customSourceIntent As Long = -1) As Boolean

    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Using embedded ICC profile to convert image from CMYK to sRGB color space..."
    #End If
    
    'Use the Color_Management module to convert the raw ICC profile into an internal Windows profile handle.  Note that
    ' this function will also validate the profile for us.
    Dim srcProfile As Long
    srcProfile = LoadICCProfileFromMemory(iccProfilePointer, iccProfileSize)
    
    'If we successfully opened and validated our source profile, continue on to the next step!
    If srcProfile <> 0 Then
    
        'Now it is time to determine our destination profile.  Because PhotoDemon operates on DIBs that default
        ' to the sRGB space, that's the profile we want to use for transformation.
            
        'Use the Color_Management module to request a standard sRGB profile.
        Dim dstProfile As Long
        dstProfile = LoadStandardICCProfile(LCS_sRGB)
        
        'It's highly unlikely that a request for a standard ICC profile will fail, but just be safe, double-check the
        ' returned handle before continuing.
        If dstProfile <> 0 Then
            
            'We can now use our profile matrix to generate a transformation object, which we will use to directly modify
            ' the DIB's RGB values.
            Dim iccTransformation As Long
            iccTransformation = RequestProfileTransform(srcProfile, dstProfile, INTENT_PERCEPTUAL, customSourceIntent)
            
            'If the transformation was generated successfully, carry on!
            If iccTransformation <> 0 Then
                
                'The only transformation function relevant to PD involves the use of BitmapBits, so we will provide
                ' the API with direct access to our DIB bits.
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "CMYK to sRGB transform data created successfully.  Applying transform..."
                #End If
                
                'Note that a color format must be explicitly specified - we vary this contingent on the parent image's
                ' color depth.
                Dim transformCheck As Boolean
                transformCheck = ApplyColorTransformToTwoDIBs(iccTransformation, srcCMYKDIB, dstRGBDIB, BM_KYMCQUADS, BM_RGBTRIPLETS)
                
                'If the transform was successful, pat ourselves on the back.
                If transformCheck Then
                    
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "CMYK to sRGB transformation successful."
                    #End If
                    
                    ApplyCMYKTransform = True
                    
                Else
                    Message "sRGB transform could not be applied.  Image remains in CMYK format."
                End If
                
                'Release our transformation
                ReleaseColorTransform iccTransformation
                                
            Else
                Message "Both ICC profiles loaded successfully, but CMYK transformation could not be created."
                ApplyCMYKTransform = False
            End If
        
            ReleaseICCProfile dstProfile
        
        Else
            Message "Could not obtain standard sRGB color profile.  CMYK transform abandoned."
            ApplyCMYKTransform = False
        End If
        
        ReleaseICCProfile srcProfile
    
    Else
        Message "Embedded ICC profile is invalid.  CMYK transform could not be performed."
        ApplyCMYKTransform = False
    End If

End Function

'When the main PD window is moved, the window manager will trigger this function.  (Because the user can set color management
' on a per-monitor basis, we must keep track of which monitor contains this PD instance.)
Public Sub CheckParentMonitor(Optional ByVal suspendRedraw As Boolean = False, Optional ByVal forceRefresh As Boolean = False)

    'Use the API to determine the monitor with the largest intersect with this window
    Dim monitorCheck As Long
    monitorCheck = MonitorFromWindow(FormMain.hWnd, MONITOR_DEFAULTTONEAREST)
    
    'If the detected monitor does not match this one, update this window and refresh its image (if necessary)
    If (monitorCheck <> currentMonitor) Or forceRefresh Then
        
        currentMonitor = monitorCheck
        currentColorProfile = g_UserPreferences.GetPref_String("Transparency", "MonitorProfile_" & currentMonitor, "")
        
        'If the user doesn't want us to redraw the main window to match the new profile, exit
        If suspendRedraw Then Exit Sub
        
        'If no images have been loaded, exit
        If pdImages(g_CurrentImage) Is Nothing Then Exit Sub
        
        'If an image has been loaded, and it is valid, redraw it now
        If (pdImages(g_CurrentImage).Width > 0) And (pdImages(g_CurrentImage).Height > 0) And (FormMain.WindowState <> vbMinimized) And (g_WindowManager.GetClientWidth(FormMain.hWnd) > 0) And pdImages(g_CurrentImage).loadedSuccessfully Then
            Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        End If
        
        'Note that the image tabstrip is also color-managed, so it needs to be redrawn as well
        toolbar_ImageTabs.forceRedraw
    
    End If
    
End Sub

'Compare two ICC profiles to determine equality.  Thank you to VB developer LaVolpe for this suggestion and original implementation.
Public Function AreColorProfilesEqual(ByVal profileHandle1 As Long, ByVal profileHandle2 As Long) As Boolean

    Dim profilesEqual As Boolean
    profilesEqual = True

    'ICC profiles headers are fixed-length (128 bytes)
    Dim firstHeader(0 To 127) As Byte, secondHeader(0 To 127) As Byte
    
    If GetColorProfileHeader(profileHandle1, VarPtr(firstHeader(0))) <> 0 Then
        If GetColorProfileHeader(profileHandle2, VarPtr(secondHeader(0))) <> 0 Then
                    
            Dim x As Long
            For x = 1 To 127
                If firstHeader(x) <> secondHeader(x) Then
                    profilesEqual = False
                    Exit For
                End If
            Next x
            
        End If
    End If
    
    AreColorProfilesEqual = profilesEqual
    
End Function

'RGB to XYZ conversion using custom endpoints requires a special transform.  We cannot use Microsoft's built-in transform methods as they do
' not support variable white space endpoints (WTF, MICROSOFT).
'
'Note that this function supports transforms in *both* directions!  The optional treatEndpointsAsForwardValues can be set to TRUE to use the
' endpoint math when converting TO the XYZ space; if false, sRGB will be used for the RGB -> XYZ conversion, then the optional parameters
' will be used for the XYZ -> RGB conversion.  Directionality is important when working with filetypes (like PNG) that specify their own
' endpoints, as the endpoints define the reverse transform, not the forward one.  (Found this out the hard way, ugh.)
'
'Gamma is also optional; if none is specified, the default sRGB gamma transform value will be used.  Note that sRGB uses a two-part curve
' constructed around 2.4 - *not* a simple one-part 2.2 curve - so if you want 2.2 gamma, make sure you specify it!
Public Function ConvertRGBUsingCustomEndpoints(ByRef srcDIB As pdDIB, ByVal RedX As Double, ByVal RedY As Double, ByVal GreenX As Double, ByVal GreenY As Double, ByVal BlueX As Double, ByVal BlueY As Double, ByVal WhiteX As Double, ByVal WhiteY As Double, Optional ByRef srcGamma As Double = 0#, Optional ByVal treatEndpointsAsForwardValues As Boolean = False, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean

    'As always, Bruce Lindbloom provides very helpful conversion functions here:
    ' http://brucelindbloom.com/index.html?Eqn_RGB_to_XYZ.html
    '
    'The biggest problem with XYZ conversion is inverting the calculation matrix.  This is a big headache, as VB provides no inherent
    ' matrix functions, so we have to do everything manually.
    
    'Start by calculating an XYZ triplet that corresponds to the incoming white point value
    Dim Xw As Double, Yw As Double, Zw As Double
    Xw = WhiteX / WhiteY
    Yw = 1
    Zw = (1 - WhiteX - WhiteY) / WhiteY
    
    'Next, calculate xyz triplets that correspond to the incoming RGB endpoints, using the same xyz to XYZ conversion as the white point.
    Dim Xr As Double, Yr As Double, Zr As Double
    Dim Xg As Double, Yg As Double, Zg As Double
    Dim Xb As Double, Yb As Double, Zb As Double
    
    Xr = RedX / RedY
    Yr = 1
    Zr = (1 - RedX - RedY) / RedY
    
    Xg = GreenX / GreenY
    Yg = 1
    Zg = (1 - GreenX - GreenY) / GreenY
    
    Xb = BlueX / BlueY
    Yb = 1
    Zb = (1 - BlueX - BlueY) / BlueY
    
    'Now comes the ugly stuff.  We can think of the calculated XYZ values (for each of RGB) as a conversion matrix, which looks like this:
    ' [Xr Xg Xb
    '  Yr Yg Yb
    '  Zr Zg Zb]
    '
    'We want to calculate a new conversion vector, [Sr Sg Sb], that takes into account the white endpoints specified above.  To calculate
    ' such a vector, we need to multiple the white point vector [Xw Yw Zw] by the *inverse* of the matrix above.  Matrix inversion is
    ' unpleasant work, as VB provides no internal function for it - so we must invert it manually.
    '
    'There are a number of different matrix inversion algorithms, but I'm going to use Gaussian elimination, as it's one of the few I
    ' remember from school.  Thanks to Vagelis Plevris of Greece, whose FreeVBCode project provided a nice refresher on how one can
    ' tackle this in VB (http://www.freevbcode.com/ShowCode.asp?ID=6221).
    Dim invMatrix() As Double, srcMatrix() As Double
    ReDim invMatrix(0 To 2, 0 To 2) As Double
    ReDim srcMatrix(0 To 2, 0 To 2) As Double
    
    srcMatrix(0, 0) = Xr
    srcMatrix(0, 1) = Xg
    srcMatrix(0, 2) = Xb
    srcMatrix(1, 0) = Yr
    srcMatrix(1, 1) = Yg
    srcMatrix(1, 2) = Yb
    srcMatrix(2, 0) = Zr
    srcMatrix(2, 1) = Zg
    srcMatrix(2, 2) = Zb
    
    'Apply the inversion.  Note that *not all matrices are invertible*!  Image-encoded endpoints should be valid, but if they are not,
    ' matrix inversion will fail.
    If Invert3x3Matrix(invMatrix, srcMatrix) Then
        
        'Calculate the S conversion vector by multiplying the inverse matrix by the white point vector
        Dim Sr As Double, Sg As Double, Sb As Double
        Sr = invMatrix(0, 0) * Xw + invMatrix(0, 1) * Yw + invMatrix(0, 2) * Zw
        Sg = invMatrix(1, 0) * Xw + invMatrix(1, 1) * Yw + invMatrix(1, 2) * Zw
        Sb = invMatrix(2, 0) * Xw + invMatrix(2, 1) * Yw + invMatrix(2, 2) * Zw
        
        'We now have everything we need to calculate the primary transformation matrix [M], which is used as follows:
        ' [X Y Z] = [M][R G B]
        Dim mFinal() As Double
        ReDim mFinal(0 To 2, 0 To 2) As Double
        mFinal(0, 0) = Sr * Xr
        mFinal(0, 1) = Sg * Xg
        mFinal(0, 2) = Sb * Xb
        mFinal(1, 0) = Sr * Yr
        mFinal(1, 1) = Sg * Yg
        mFinal(1, 2) = Sb * Yb
        mFinal(2, 0) = Sr * Zr
        mFinal(2, 1) = Sg * Zg
        mFinal(2, 2) = Sb * Zb
        
        'Debug.Print "Forward matrix: "
        'Debug.Print mFinal(0, 0), mFinal(0, 1), mFinal(0, 2)
        'Debug.Print mFinal(1, 0), mFinal(1, 1), mFinal(1, 2)
        'Debug.Print mFinal(2, 0), mFinal(2, 1), mFinal(2, 2)
        
        'Want to convert from XYZ to RGB?  Use the inverse matrix!  This is required for PNG files, because their endpoints specify
        ' the reverse transform.  I'm not sure why this is.  My matrix math is rusty, but it's possible that we could skip the
        ' first inversion, and simply multiply the S vector to the original source matrix, but I haven't tried this to see if it works
        ' and my math skills are too rusty to know if that's a totally invalid operation.  As such, I just invert the matrix manually,
        ' to be safe.
        '
        'Note that we don't have to check for a fail state here, as we know the matrix is invertible (because we inverted it ourselves
        ' earlier on.  It's technically possible for faulty white point values to prevent this inversion, but PD doesn't provide a way
        ' for users to enter faulty values, so I don't check that possibility here.
        Dim mFinalInvert() As Double
        ReDim mFinalInvert(0 To 2, 0 To 2) As Double
        If Not treatEndpointsAsForwardValues Then Invert3x3Matrix mFinalInvert, mFinal
        
        'Debug.Print "Reverse matrix: "
        'Debug.Print mFinalInvert(0, 0), mFinalInvert(0, 1), mFinalInvert(0, 2)
        'Debug.Print mFinalInvert(1, 0), mFinalInvert(1, 1), mFinalInvert(1, 2)
        'Debug.Print mFinalInvert(2, 0), mFinalInvert(2, 1), mFinalInvert(2, 2)
        
        'We now have everything we need to convert the DIB.  PARTY TIME!
        Dim x As Long, y As Long
        
        'The actual XYZ transform is actually pretty simple.  Using the supplied endpoints, we use our custom matrix either during the
        ' RGB -> XYZ step (forward transform), or XYZ -> RGB step (reverse transform).  The unused stage uses hard-coded sRGB values,
        ' including hard-coded sRGB gamma (unless another gamma was specified).
        '
        'For forward transforms, we must pre-linearize the RGB values, using the supplied gamma.  Because the gamma is applied to the
        ' [0, 255] range RGB values, we can use a look-up table to accelerate the process.
        Dim gammaLookup(0 To 255) As Double, tmpCalc As Double
        If treatEndpointsAsForwardValues Then
            
            'Invert gamma if it was specified
            If srcGamma <> 0 Then srcGamma = 1 / srcGamma
            
            For x = 0 To 255
                tmpCalc = x / 255
                
                If srcGamma = 0 Then
                    
                    If tmpCalc > 0.04045 Then
                        gammaLookup(x) = ((tmpCalc + 0.055) / (1.055)) ^ 2.4
                    Else
                        gammaLookup(x) = tmpCalc / 12.92
                    End If
                    
                Else
                    gammaLookup(x) = tmpCalc ^ srcGamma
                End If
                
            Next x
        
        'Reverse transforms apply gamma directly to the floating-point RGB values, to reduce data loss due to clamping.
        Else
        
            'Do nothing for reverse transforms, except to invert gamma as appropriate.
            If srcGamma <> 0 Then srcGamma = 1 / srcGamma
            
        End If
        
        Dim tmpX As Double, tmpY As Double, tmpZ As Double
        
        'Create a local array and point it at the pixel data we want to operate on
        Dim ImageData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
            
        'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
        Dim initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.getDIBWidth - 1
        finalY = srcDIB.getDIBHeight - 1
                
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim QuickVal As Long, qvDepth As Long
        qvDepth = srcDIB.getDIBColorDepth \ 8
        
        'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
        ' based on the size of the area to be processed.
        Dim progBarCheck As Long
        If Not suppressMessages Then
            If modifyProgBarMax = -1 Then
                SetProgBarMax finalX
            Else
                SetProgBarMax modifyProgBarMax
            End If
            progBarCheck = findBestProgBarValue()
        End If
        
        'Color values
        Dim r As Long, g As Long, b As Long
        Dim fR As Double, fG As Double, fB As Double
                
        'Now we can loop through each pixel in the image, converting values as we go
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
                
            'Get the source pixel color values
            r = ImageData(QuickVal + 2, y)
            g = ImageData(QuickVal + 1, y)
            b = ImageData(QuickVal, y)
            
            'Branch now according to forward/reverse transforms.
            
            'Forward transform
            If treatEndpointsAsForwardValues Then
            
                'Convert to compressed gamma representation
                fR = gammaLookup(r)
                fG = gammaLookup(g)
                fB = gammaLookup(b)
                
                'Convert to XYZ
                tmpX = mFinal(0, 0) * fR + mFinal(0, 1) * fG + mFinal(0, 2) * fB
                tmpY = mFinal(1, 0) * fR + mFinal(1, 1) * fG + mFinal(1, 2) * fB
                tmpZ = mFinal(2, 0) * fR + mFinal(2, 1) * fG + mFinal(2, 2) * fB
                
                'Convert back to sRGB
                Color_Functions.XYZtoRGB tmpX, tmpY, tmpZ, r, g, b
            
            'Reverse transform
            Else
            
                'Use sRGB for the initial XYZ conversion
                Color_Functions.RGBtoXYZ r, g, b, tmpX, tmpY, tmpZ
            
                'Convert back to [0, 1] RGB, using our custom endpoints
                fR = mFinalInvert(0, 0) * tmpX + mFinalInvert(0, 1) * tmpY + mFinalInvert(0, 2) * tmpZ
                fG = mFinalInvert(1, 0) * tmpX + mFinalInvert(1, 1) * tmpY + mFinalInvert(1, 2) * tmpZ
                fB = mFinalInvert(2, 0) * tmpX + mFinalInvert(2, 1) * tmpY + mFinalInvert(2, 2) * tmpZ
            
                'Convert to linear RGB, accounting for gamma
                If srcGamma = 0 Then
                
                    'If the user didn't specify gamma, use a default sRGB transform.
                    If (fR > 0.0031308) Then
                        fR = 1.055 * (fR ^ (1 / 2.4)) - 0.055
                    Else
                        fR = 12.92 * fR
                    End If
                    
                    If (fG > 0.0031308) Then
                        fG = 1.055 * (fG ^ (1 / 2.4)) - 0.055
                    Else
                        fG = 12.92 * fG
                    End If
                    
                    If (fB > 0.0031308) Then
                        fB = 1.055 * (fB ^ (1 / 2.4)) - 0.055
                    Else
                        fB = 12.92 * fB
                    End If
                
                Else
                
                    If fR > 0 Then
                        r = (fR ^ srcGamma) * 255
                    Else
                        r = 0
                    End If
                    
                    If fG > 0 Then
                        g = (fG ^ srcGamma) * 255
                    Else
                        g = 0
                    End If
                    
                    If fB > 0 Then
                        b = (fB ^ srcGamma) * 255
                    Else
                        b = 0
                    End If
                
                End If
                
                'Apply RGB clamping now
                If r > 255 Then
                    r = 255
                ElseIf r < 0 Then
                    r = 0
                End If
                
                If g > 255 Then
                    g = 255
                ElseIf g < 0 Then
                    g = 0
                End If
                
                If b > 255 Then
                    b = 255
                ElseIf b < 0 Then
                    b = 0
                End If
                
            End If
            
            'Assign the new colors and continue
            ImageData(QuickVal, y) = b
            ImageData(QuickVal + 1, y) = g
            ImageData(QuickVal + 2, y) = r
            
        Next y
            If Not suppressMessages Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal x + modifyProgBarOffset
                End If
            End If
        Next x
                
        'With our work complete, point ImageData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
        Erase ImageData
        
        If cancelCurrentAction Then ConvertRGBUsingCustomEndpoints = False Else ConvertRGBUsingCustomEndpoints = True
        
    Else
        ConvertRGBUsingCustomEndpoints = False
        Exit Function
    End If
    
End Function

'Invert a 3x3 matrix of double-type.  The matrices MUST BE DIMMED PROPERLY PRIOR TO CALLING THIS FUNCTION.
' Failure returns FALSE; success, TRUE.
'
'Thanks to Vagelis Plevris of Greece, whose FreeVBCode project provided a nice refresher on how one might
' tackle this in VB (http://www.freevbcode.com/ShowCode.asp?ID=6221).
Private Function Invert3x3Matrix(ByRef newMatrix() As Double, ByRef srcMatrix() As Double) As Boolean

    'Some matrices are not invertible.  If color endpoints are calculated correctly, this shouldn't be a problem,
    ' but we need to have a failsafe for the case of determinant = 0
    On Error GoTo cantCreateMatrix
    
    'Gaussian elimination will use an intermediate array, at double the width of the incoming srcMatrix()
    Dim intMatrix() As Double
    ReDim intMatrix(0 To 2, 0 To 5) As Double
    
    'Gaussian elimination works by using simple row operations to solve a system of linear equations.  This is
    ' computationally slow, but algorithmically simple, and for a single 3x3 matrix no one cares about performance.
    '
    'To visualize what happens, see how we put the source matrix on the left and the identity matrix on the right, like so:
    '
    ' [ src11 src12 src13 | 1 0 0 ]
    ' [ src21 src22 src23 | 0 1 0 ]
    ' [ src31 src32 src33 | 0 0 1 ]
    '
    'When we're done, we will have constructed the inverse on the right, as a result of our row operations:
    ' [ 1 0 0 | inv11 inv12 inv13 ]
    ' [ 0 1 0 | inv21 inv22 inv23 ]
    ' [ 0 0 1 | inv31 inv32 inv33 ]
    
    'Start by filling our calculation array with the input values
    Dim x As Long, y As Long
    For x = 0 To 2
    For y = 0 To 2
        intMatrix(x, y) = srcMatrix(x, y)
    Next y
    Next x
    
    'Populate the identity matrix on the right
    intMatrix(0, 3) = 1
    intMatrix(1, 4) = 1
    intMatrix(2, 5) = 1
    
    'Start performing row operations that move us toward an identity matrix on the left
    Dim k As Long, n As Long, m As Long, nonZeroLine As Long, tmpValue As Double
    
    For k = 0 To 2
        
        'A non-zero element is required.  Change lines if necessary to make this happen.
        If intMatrix(k, k) = 0 Then
            
            'Find the first line with a non-zero element
            For n = k To 2
                If intMatrix(n, k) <> 0 Then
                    nonZeroLine = n
                    Exit For
                End If
            Next n
            
            'Swap line k and nonZeroLine
            For m = k To 5
                tmpValue = intMatrix(k, m)
                intMatrix(k, m) = intMatrix(nonZeroLine, m)
                intMatrix(nonZeroLine, m) = tmpValue
            Next m
            
        End If
            
        tmpValue = intMatrix(k, k)
        For n = k To 5
            intMatrix(k, n) = intMatrix(k, n) / tmpValue
        Next n
        
        'For other lines, make a zero element using the formula:
        ' Ai1 = Aij - A11 * (Aij / A11)
        For n = 0 To 2
            
            'Check finishing position
            If (n = k) And (n = 2) Then Exit For
            
            'Check for elements already equal to one; it's not really good form to update a loop element like this,
            ' but it's helpful in the absence of an easy way to tell VB to "Goto Next"
            If (n = k) And (n < 2) Then n = n + 1
            
            'Do not touch elements that are already zero
            If intMatrix(n, k) <> 0 Then
            
                If intMatrix(k, k) <> 0 Then
                    
                    tmpValue = intMatrix(n, k) / intMatrix(k, k)
                    For m = k To 5
                        intMatrix(n, m) = intMatrix(n, m) - intMatrix(k, m) * tmpValue
                    Next m
                    
                'Failed determinant; exit function
                Else
                
                    GoTo cantCreateMatrix
                
                End If
                
            End If
            
        Next n
        
    Next k
    
    'Inversion complete!  (Barring any divide-by-zero errors, which indicate an un-invertible matrix.)
    
    'Copy the solved section of the intermediate matrix into the destination
    For n = 0 To 2
    For k = 0 To 2
        newMatrix(n, k) = intMatrix(n, 3 + k)
    Next k
    Next n
    
    'Report the successful inversion to the user, then exit
    Invert3x3Matrix = True
    Exit Function

cantCreateMatrix:
    Debug.Print "Matrix is not invertible; function cancelled."
    Invert3x3Matrix = False

End Function
