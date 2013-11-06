Attribute VB_Name = "Color_Management"
'***************************************************************************
'PhotoDemon ICC (International Color Consortium) Profile Support Module
'Copyright ©2012-2013 by Tanner Helland
'Created: 05/November/13
'Last updated: 05/November/13
'Last update: moved some code elements out of the pdICCProfile class and into this standalone support module.
'              Because PD intends to color manage multiple parts of the interface (not just raw image data, but
'              also picture boxes and screens being rendered to), it proved useful to build some standardized
'              ICC-related functions that can be reused under various circumstances.
'
'ICC profiles can be embedded in certain types of images (JPEG, PNG, and TIFF at the time of this writing).  These
' profiles can be used to convert an image to its true color space, taking into account any pecularities of the
' device that captured the image (typically a camera), and the device now being used to display the image
' (typically a monitor).
'
'ICC profile handling is broken into two parts: extracting the profile from an image, then applying that profile
' to the image.  The extraction step is currently handled via FreeImage or GDI+, while the application step is handled
' by Windows.  In the future I may look at adding ExifTool as a possibly mechanism for extracting the profile, as it
' provides better support for esoteric formats than FreeImage.
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
Public Const LCS_sRGB As Long = &H73524742
Public Const LCS_WINDOWS_COLOR_SPACE As Long = &H2

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
Private Const BEST_MODE As Long = 3&

'Because we only do ICC-to-ICC transforms, Windows can be instructed to use a 3rd-party CMS instead of its own
' internal one.  We don't care if it does this, and we tell it as much.
Private Const INDEX_DONT_CARE As Long = 0&

'When it comes time to actually apply the transformation to the image data, the transform needs to know the image's
' color depth.  PD only operates in 24 and 32bpp mode, so we only need two constants here.
Private Const BM_RGBTRIPLETS As Long = &H2
Private Const BM_BGRTRIPLETS As Long = &H4
Private Const BM_xBGRQUADS As Long = &H10

'Various ICC-related APIs are needed to open profiles and transform data between them
Private Declare Function OpenColorProfile Lib "mscms" Alias "OpenColorProfileA" (ByRef pProfile As Any, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal dwCreationMode As Long) As Long
Private Declare Function CloseColorProfile Lib "mscms" (ByVal hProfile As Long) As Long
Private Declare Function IsColorProfileValid Lib "mscms" (ByVal hProfile As Long, ByRef pBool As Long) As Long
Private Declare Function GetStandardColorSpaceProfile Lib "mscms" Alias "GetStandardColorSpaceProfileA" (ByVal pcstr As String, ByVal dwProfileID As Long, ByVal pProfileName As Long, ByRef pdwSize As Long) As Long
Private Declare Function CreateMultiProfileTransform Lib "mscms" (ByRef pProfile As Any, ByVal nProfiles As Long, ByRef pIntents As Long, ByVal nIntents As Long, ByVal dwFlags As Long, ByVal indexPreferredCMM As Long) As Long
Private Declare Function DeleteColorTransform Lib "mscms" (ByVal hTransform As Long) As Long
Private Declare Function TranslateBitmapBits Lib "mscms" (ByVal hTransform As Long, ByRef srcBits As Any, ByVal pBmInput As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal dwInputStride As Long, ByRef dstBits As Any, ByVal pBmOutput As Long, ByVal dwOutputStride As Long, ByRef pfnCallback As Long, ByVal ulCallbackData As Long) As Long

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
Private Const MAX_PATH As Long = 260

'Assign the default color profile (whether the system profile or the user profile) to a DC
Public Sub assignDefaultColorProfileToDC(ByVal targetDC As Long)
    SetICMProfile targetDC, currentSystemColorProfile
    
    'If you would like to test this function on a standalone ICC profile (generally something bizarre, to help you know
    ' that the function is working), use something similar to the code below.
    'Dim TEST_ICM As String
    'TEST_ICM = "C:\PhotoDemon v4\PhotoDemon\no_sync\Images from testers\jpegs\ICC\WhackedRGB.icc"
    'SetICMProfile targetDC, TEST_ICM
End Sub

'When PD is first loaded, this function will be called, which caches the current color management file in use by the system
Public Sub cacheCurrentSystemColorProfile()
    currentSystemColorProfile = getDefaultICCProfile()
End Sub

'Returns the path to the default color mangement profile file (ICC or WCS) currently in use by the system.
Public Function getDefaultICCProfile() As String

    'Prepare a blank string to receive the profile path
    Dim filenameLength As Long
    filenameLength = MAX_PATH
    
    Dim tmpPathString As String
    tmpPathString = ""
    
    Dim tmpPathBuffer() As Byte
    ReDim tmpPathBuffer(0 To filenameLength - 1) As Byte
    
    'Using the desktop DC as our reference, request the filename of the currently in-use ICM profile (which should be the system default)
    If GetICMProfile(GetDC(0), filenameLength, ByVal VarPtr(tmpPathBuffer(0))) = 0 Then
        getDefaultICCProfile = ""
    Else
    
        'Convert the returned byte array into a string
        Dim i As Long
        For i = 0 To filenameLength - 1
            tmpPathString = tmpPathString & Chr(tmpPathBuffer(i))
        Next i
                
        getDefaultICCProfile = tmpPathString
        
    End If
    
End Function

'Turn on color management for a specified device context
Public Sub turnOnColorManagementForDC(ByVal dstDC As Long)
    SetICMMode dstDC, ICM_ON
End Sub

'Turn off color management for a specified device context
Public Sub turnOffColorManagementForDC(ByVal dstDC As Long)
    SetICMMode dstDC, ICM_OFF
End Sub

'Given a valid iccProfileArray (such as one stored in a pdICCProfile class), convert it to an internal Windows color profile
' handle, validate it, and return the result.  Returns a non-zero handle if successful.
Public Function loadICCProfileFromMemory(ByVal profileArrayPointer As Long, ByVal profileArraySize As Long) As Long

    'Start by preparing an ICC_PROFILE header to use with the color management APIs
    Dim srcProfileHeader As ICC_PROFILE
    srcProfileHeader.dwType = PROFILE_MEMBUFFER
    srcProfileHeader.pProfileData = profileArrayPointer
    srcProfileHeader.cbDataSize = profileArraySize
    
    'Use that header to open a reference to an internal Windows color profile (which is required by all ICC-related API)
    loadICCProfileFromMemory = OpenColorProfile(srcProfileHeader, PROFILE_READ, FILE_SHARE_READ, OPEN_EXISTING)
    
    If loadICCProfileFromMemory <> 0 Then
    
        'Validate the profile's XML as well; it is possible for a profile to be ill-formed, which means we cannot use it.
        Dim tmpCheck As Long
        If IsColorProfileValid(loadICCProfileFromMemory, tmpCheck) = 0 Then
            Debug.Print "Color profile loaded succesfully, but XML failed to validate."
            CloseColorProfile loadICCProfileFromMemory
            loadICCProfileFromMemory = 0
        End If
        
    Else
        Debug.Print "ICC profile failed to load (OpenColorProfile failed with error #" & Err.LastDllError & ")."
    End If

End Function

'Request a standard ICC profile from the OS.  Windows only provides two standard color profiles: sRGB (LCS_sRGB), and whatever
' the system default currently is (LCS_WINDOWS_COLOR_SPACE).  While probably not necessary, this function also validates the
' requested profile, just to be safe.
Public Function loadStandardICCProfile(ByVal profileID As Long) As Long

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
    Dim dstProfile As Long
    loadStandardICCProfile = OpenColorProfile(dstProfileHeader, PROFILE_READ, FILE_SHARE_READ, OPEN_EXISTING)
    
    'It is highly unlikely (maybe even impossible?) for the system to return an invalid standard profile, but just to be
    ' safe, validate the XML.
    If loadStandardICCProfile <> 0 Then
    
        'Validate the profile's XML as well; it is possible for a profile to be ill-formed, which means we cannot use it.
        Dim tmpCheck As Long
        If IsColorProfileValid(loadStandardICCProfile, tmpCheck) = 0 Then
            Debug.Print "Standard color profile loaded succesfully, but XML failed to validate."
            CloseColorProfile loadStandardICCProfile
            loadStandardICCProfile = 0
        End If
        
    Else
        Debug.Print "Standard ICC profile failed to load (GetStandardColorSpaceProfile failed with error #" & Err.LastDllError & ")."
    End If

End Function

'This function is just a thin wrapper to CloseColorProfile; however, using it allows us to keep various color-management
' DLLs nicely encapsulated within this module.
Public Sub releaseICCProfile(ByVal profileHandle As Long)
    CloseColorProfile profileHandle
End Sub

'Given a source profile, destination profile, and rendering intent, return a compatible transformation handle.
Public Function requestProfileTransform(ByVal srcProfile As Long, ByVal dstProfile As Long, ByVal preferredIntent As RenderingIntents) As Long

     'Next we need to prepare two matrices to supply to CreateMultiProfileTransform: one for ICC profiles themselves,
    ' and one for desired render intents.
    Dim profileMatrix(0 To 1) As Long, intentMatrix(0 To 1) As Long
    
    'The first row in the array contains the two profile pointers we've already acquired, in src/dest order
    profileMatrix(0) = srcProfile
    profileMatrix(1) = dstProfile
    
    'The second column in the array contains the render intents for the transformation.
    intentMatrix(0) = preferredIntent
    intentMatrix(1) = preferredIntent
    
    'We can now use our profile matrix to generate a transformation object, which we will use on the DIB itself
    requestProfileTransform = CreateMultiProfileTransform(ByVal VarPtr(profileMatrix(0)), 2&, ByVal VarPtr(intentMatrix(0)), 2&, BEST_MODE, INDEX_DONT_CARE)
    
    If requestProfileTransform = 0 Then
        Debug.Print "Requested color transformation could not be generated (Error #" & Err.LastDllError & ")."
    End If
    
End Function

'This function is just a thin wrapper to DeleteColorTransform; however, using it allows us to keep various color-management
' DLLs nicely encapsulated within this module.
Public Sub releaseColorTransform(ByVal transformHandle As Long)
    DeleteColorTransform transformHandle
End Sub

'Given a color transformation and a layer, apply one to the other!  Returns TRUE if successful.
Public Function applyColorTransformToLayer(ByVal srcTransform As Long, ByRef dstLayer As pdLayer) As Boolean

    Dim transformCheck As Long
    
    With dstLayer
                
        'NOTE: note that I use BM_RGBTRIPLETS below, despite pdLayer DIBs most definitely being in BGR order.  This is an
        '       undocumented bug with Windows' color management engine!
        Dim bitDepthIdentifier As Long
        If .getLayerColorDepth = 24 Then bitDepthIdentifier = BM_RGBTRIPLETS Else bitDepthIdentifier = BM_xBGRQUADS
        
        'TranslateBitmapBits handles the actual transformation for us.
        transformCheck = TranslateBitmapBits(srcTransform, ByVal .getLayerDIBits, bitDepthIdentifier, .getLayerWidth, .getLayerHeight, .getLayerArrayWidth, ByVal .getLayerDIBits, bitDepthIdentifier, .getLayerArrayWidth, ByVal 0&, 0&)
        
    End With
    
    If transformCheck = 0 Then
        applyColorTransformToLayer = False
        
        'Error #2021 is ERROR_COLORSPACE_MISMATCH: "The specified transform does not match the bitmap's color space."
        ' This is a known error when the source image was in CMYK format, because FreeImage (or GDI+) will have
        ' automatically converted the image to RGB at load-time.  Because the ICC profile is CMYK-specific, Windows will
        ' not be able to apply it to the image, as it is no longer in CMYK format!
        If CLng(Err.LastDllError) = 2021 Then
            Debug.Print "ICC profile could not be applied, because CMYK conversion already occurred."
        Else
            Debug.Print "ICC profile could not be applied.  Image remains in original profile. (Error #" & Err.LastDllError & ")."
        End If
        
    Else
        applyColorTransformToLayer = True
    End If

End Function
