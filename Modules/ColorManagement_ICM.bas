Attribute VB_Name = "ColorManagement"
'***************************************************************************
'PhotoDemon ICC (International Color Consortium) Profile Support Module
'Copyright 2013-2016 by Tanner Helland
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
' CheckParentMonitor() function, below.
Private m_CurrentMonitor As Long

'When the main form's monitor changes, this string will automatically be updated with the corresponding ICC
' profile path of that monitor (if the user has selected a custom one)
Private m_CurrentDisplayProfile As String

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
    RI_INTENT_PERCEPTUAL = 0&
    RI_INTENT_RELATIVECOLORIMETRIC = 1&
    RI_INTENT_SATURATION = 2&
    RI_INTENT_ABSOLUTECOLORIMETRIC = 3&
End Enum

#If False Then
    Const RI_INTENT_PERCEPTUAL = 0&, RI_INTENT_RELATIVECOLORIMETRIC = 1&, RI_INTENT_SATURATION = 2&, RI_INTENT_ABSOLUTECOLORIMETRIC = 3&
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
Private Declare Function GetColorDirectory Lib "mscms" Alias "GetColorDirectoryW" (ByVal pMachineName As Long, ByVal ptrToBuffer As Long, ByRef pdwSize As Long) As Long
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
Private Declare Function GetICMProfile Lib "gdi32" Alias "GetICMProfileW" (ByVal hDC As Long, ByRef lpcbName As Long, ByVal ptrToBuffer As Long) As Long
Private Declare Function SetICMProfile Lib "gdi32" Alias "SetICMProfileW" (ByVal hDC As Long, ByVal ptrToFilename As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

'When PD is first loaded, the system's current color management file will be cached in this variable
Private m_currentSystemColorProfile As String
Private Const MAX_PATH As Long = 260

'PD supports several different display color management modes
Public Enum DISPLAY_COLOR_MANAGEMENT
    DCM_NoManagement = 0
    DCM_SystemProfile = 1
    DCM_CustomProfile = 2
End Enum

#If False Then
    Private Const DCM_NoManagement = 0, DCM_SystemProfile = 1, DCM_CustomProfile = 2
#End If

'The current display management mode, cached.  (This value is generally equal to the return value of
' GetDisplayColorManagementPreference(), but - for example - if the user has selected "use system color profile"
' and Windows doesn't return a valid profile, we will automatically switch to unmanaged mode.)
Private m_DisplayCMMPolicy As DISPLAY_COLOR_MANAGEMENT

'Current display render intent.  As of 7.0, the user can change this from Perceptual (although it's really not
' recommended unless soft-proofing is active).
Private m_DisplayRenderIntent As LCMS_RENDERING_INTENT

'PD supports several different layer color management modes
Public Enum LAYER_COLOR_MANAGEMENT
    LCM_NoManagement = 0
    LCM_ProfileTagged = 1
    LCM_ProfileConverted = 2
End Enum

#If False Then
    Private Const LCM_NoManagement = 0, LCM_ProfileTagged = 1, LCM_ProfileConverted = 2
#End If

'If loaded successfully, the current system profile will be found at this index into the profile cache.  Note that
' this value *does not* support multiple monitors, due to flaws in the way Windows exposes the system profile
' (you can only easily retrieve the default monitor profile).  As such, if the user has custom-configured a
' monitor profile in the Tools > Options dialog, you shouldn't be using this value at all.
Private m_SystemProfileIndex As Long

'Current display index.  This value is automatically refreshed by calls to CheckParentMonitor, below, and is used
' to support multimonitor systems.  On some configurations, it may be identical to m_SystemProfileIndex, above, or
' m_sRGBIndex, below.
Private m_CurrentDisplayIndex As Long

'sRGB profile index.  One valid sRGB profile is always loaded into memory, and it is used as a failsafe when things
' go horribly wrong (e.g. no system profile is configured, a requested display profile is corrupted or missing, etc).
Private m_sRGBIndex As Long

'Cached profiles contain more information than a typical ICC profile, to make it easier to match profiles
' against each other.
Private Type ICCProfileCache
    
    'All profiles must be accompanied of a bytestream copy of their ICC profile contents.  This allows us to do things like
    ' match duplicates, or look for duplicate source paths.
    fullProfile As pdICCProfile
    
    'Generally speaking, all profiles should also be accompanied by a LittleCMS handle to said profile.  Note that these
    ' handles will leak if not manually released!
    lcmsProfileHandle As Long
    
    'Flags to help us shortcut searches for certain profile types
    isSystemProfile As Boolean
    isPDDisplayProfile As Boolean
    isWorkingSpaceProfile As Boolean
    
    'This field is only used for display profiles; it is the HMONITOR corresponding to this display (used to match profiles
    ' to displays at run-time)
    curDisplayID As Long
    
    'When used as part of the main viewport pipeline, working-space profiles cache a reference to a LittleCMS transform that
    ' translates between that working space and the current display profile.  This transform is optimized at creation time,
    ' and subsequently shared between all users of that working-space profile.  (The index of the matching display transform
    ' is also cached, so that we can detect future mismatches.)
    '
    'Note that 24-bpp and 32-bpp transforms are stored and optimized separately.  Both must be freed manually, if available.
    thisWSToDisplayTransform24 As Long
    thisWSToDisplayTransform32 As Long
    indexOfDisplayTransform As Long
    
    'A unique string ID for this profile.  This is built by taking header information from the profile, and concatenating it
    ' together into something highly unique.  This string should not be presented to the user.  It *is* valid across sessions,
    ' and is used (for example) by pdLayer objects to uniquely identify their profiles when saved to file.
    profileStringID As String
    
End Type

'Current ICC Profile cache.  As of 7.0, profiles are cached here, in this sub, and any object that uses a profile
' simply receives an index into our cache.  This lets us reuse profiles for multiple objects.
Private Const INITIAL_PROFILE_CACHE_SIZE As Long = 8
Private m_NumOfCachedProfiles As Long
Private m_ProfileCache() As ICCProfileCache

'Some transforms require us to do pre- and post-conversion alpha management.  That management is tracked here, to prevent individual
' functions from needing to track it.
Private m_PreAlphaManagementRequired As Boolean

Public Function GetDisplayColorManagementPreference() As DISPLAY_COLOR_MANAGEMENT
    GetDisplayColorManagementPreference = g_UserPreferences.GetPref_Long("ColorManagement", "Display CM Mode", DCM_NoManagement)
    
    'Past PD versions used a true/false system to control this setting.  The old setting will be "-1" if the system
    ' color profile is in use.
    If (GetDisplayColorManagementPreference < 0) Then GetDisplayColorManagementPreference = DCM_SystemProfile
End Function

Public Sub SetDisplayColorManagementPreference(ByVal newPref As DISPLAY_COLOR_MANAGEMENT)
    g_UserPreferences.SetPref_Long "ColorManagement", "Display CM Mode", newPref
End Sub

Public Function GetDisplayRenderingIntentPref() As LCMS_RENDERING_INTENT
    GetDisplayRenderingIntentPref = g_UserPreferences.GetPref_Long("ColorManagement", "Display Rendering Intent", INTENT_PERCEPTUAL)
End Function

Public Sub SetDisplayRenderingIntentPref(Optional ByVal newPref As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL)
    g_UserPreferences.SetPref_Long "ColorManagement", "Display Rendering Intent", newPref
End Sub

Public Function GetSRGBProfileIndex() As Long
    GetSRGBProfileIndex = m_sRGBIndex
End Function

'Whenever color management settings change (or at program initialization), call this function to cache transforms
' for the current screen monitor space.  Up-to-date preferences will be pulled from the user's pref file, so this
' function is *not* particularly fast.  Do not use it inside a rendering chain, for example.
Public Sub CacheDisplayCMMData()
    
    'Start by releasing the current profile collection, if one exists
    FreeProfileCache
    
    Dim tmpProfile As pdICCProfile
    Set tmpProfile = New pdICCProfile
    
    'Before doing anything else, load a default sRGB profile.  This will be used as a fallback for displays with broken
    ' color management configurations.
    
    'Use LittleCMS to generate a default sRGB profile
    Dim tmpHProfile As Long
    tmpHProfile = LittleCMS.LCMS_LoadStockSRGBProfile()
    Dim tmpByteArray() As Byte
    
    If LittleCMS.LCMS_SaveProfileToArray(tmpHProfile, tmpByteArray) Then
    
        'Wrap a dummy ICC profile object around the sRGB profile, then cache that locally
        If tmpProfile.LoadICCFromPtr(UBound(tmpByteArray) + 1, VarPtr(tmpByteArray(0))) Then
            m_sRGBIndex = AddProfileToCache(tmpProfile, False, False, False, False, True)
            m_ProfileCache(m_sRGBIndex).lcmsProfileHandle = tmpHProfile
        Else
            If (tmpHProfile <> 0) Then LittleCMS.LCMS_CloseProfileHandle tmpHProfile
            m_SystemProfileIndex = -1
            m_sRGBIndex = -1
            m_DisplayCMMPolicy = DCM_NoManagement
        End If
    
    'If we failed to generate a default sRGB profile, turn off display color management completely, because something
    ' is horribly wrong with LittleCMS.
    Else
        If (tmpHProfile <> 0) Then LittleCMS.LCMS_CloseProfileHandle tmpHProfile
        m_SystemProfileIndex = -1
        m_sRGBIndex = -1
        m_DisplayCMMPolicy = DCM_NoManagement
    End If
    
    'Note that if our sRGB profile was created successfully, its LCMS handle is not freed here.  The handle is valid for the
    ' life of this PD session, and it is not released until deep in the PD unload process.
    
    'Next, load one or more display profiles based on the user's current color management settings.
    ' (In some cases, we'll retrieve profile paths from Windows itself, while in others, we'll retrieve them
    '  from internal PD settings the user has configured.)
    m_DisplayCMMPolicy = GetDisplayColorManagementPreference()
    m_DisplayRenderIntent = GetDisplayRenderingIntentPref()
    m_SystemProfileIndex = m_sRGBIndex
    
    'Start with the case where the user just wants us to use the default Windows system ICC profile.
    If (m_DisplayCMMPolicy = DCM_SystemProfile) Then
    
        m_currentSystemColorProfile = GetDefaultICCProfilePath()
        
        Set tmpProfile = New pdICCProfile
        If tmpProfile.LoadICCFromFile(m_currentSystemColorProfile) Then
            m_SystemProfileIndex = AddProfileToCache(tmpProfile, False, True)
            
            'Create an LCMS-compatible profile handle to match
            With m_ProfileCache(m_SystemProfileIndex)
                .lcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(.fullProfile.GetICCDataPointer, .fullProfile.GetICCDataSize)
            End With
            
        'If we fail to load the current system profile, there's really no good option for continuing.  The least
        ' of many evils is to simply point the current system profile at a default sRGB profile and hope for the best.
        Else
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "System ICC profile couldn't be loaded.  Reverting to default sRGB policy..."
            #End If
            
            m_SystemProfileIndex = m_sRGBIndex
            
        End If
        
    'If the user has manually specified per-monitor display profiles, we want to cache each display profile in turn.
    ElseIf (m_DisplayCMMPolicy = DCM_CustomProfile) Then
        
        Dim tmpXML As pdXML
        Set tmpXML = New pdXML
        Dim i As Long, uniqueMonitorID As String, monICCPath As String
        Dim profileLoadedSuccessfully As Boolean, profileindex As Long
        
        For i = 0 To g_Displays.GetDisplayCount - 1
            
            profileLoadedSuccessfully = False
            
            With g_Displays.Displays(i)
                
                'Retrieve a unique ID for this display
                uniqueMonitorID = .GetUniqueDescriptor
                
                'Make it XML safe and look for a matching tag
                uniqueMonitorID = tmpXML.GetXMLSafeTagName(uniqueMonitorID)
                monICCPath = g_UserPreferences.GetPref_String("ColorManagement", "DisplayProfile_" & uniqueMonitorID, vbNullString)
                
                'If an ICC path exists, attempt to load it
                If (Len(monICCPath) <> 0) Then
                    
                    Set tmpProfile = New pdICCProfile
                    If tmpProfile.LoadICCFromFile(monICCPath) Then
                        
                        'Add the profile to our collection!
                        profileindex = AddProfileToCache(tmpProfile, False, False, True, .GetHandle)
                        
                        'Create an LCMS-compatible profile handle to match
                        If (profileindex >= 0) Then
                            With m_ProfileCache(profileindex)
                                .lcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(.fullProfile.GetICCDataPointer, .fullProfile.GetICCDataSize)
                            End With
                            profileLoadedSuccessfully = True
                        End If
                        
                    End If
                
                End If
                
                'If a profile was *not* loaded successfully, default to sRGB for this display
                If (Not profileLoadedSuccessfully) Then
                    profileindex = AddProfileToCache(GetCachedProfile_ByIndex(m_sRGBIndex).fullProfile, False, False, True, .GetHandle)
                    
                    If (profileindex >= 0) Then
                        With m_ProfileCache(profileindex)
                            .lcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(.fullProfile.GetICCDataPointer, .fullProfile.GetICCDataSize)
                        End With
                    End If
                End If
                
            End With
        Next i
        
        'With all display profiles loaded correctly, identify the current display immediately
        CheckParentMonitor False, True
    
    'If the user has selected "no color management", we don't have to do anything here!
    Else
    
    End If
    
End Sub

'Add a profile to the current cache.  The index of said profile is returned; use that for any subsequent cache accesses.
Public Function AddProfileToCache(ByRef srcProfile As pdICCProfile, Optional ByVal matchDuplicates As Boolean = True, Optional ByVal isSystemProfile As Boolean = False, Optional ByVal isDisplayProfile As Boolean = False, Optional ByVal associatedMonitorID As Long = 0, Optional ByVal isWorkingSpace As Boolean = False) As Long
    
    'Make sure the cache exists and is large enough to hold another profile
    If (m_NumOfCachedProfiles = 0) Then
        ReDim m_ProfileCache(0 To INITIAL_PROFILE_CACHE_SIZE - 1) As ICCProfileCache
    Else
        If (m_NumOfCachedProfiles > UBound(m_ProfileCache)) Then ReDim m_ProfileCache(0 To (UBound(m_ProfileCache) * 2 + 1)) As ICCProfileCache
    End If
    
    'If the user wants profile matching to occur (so that duplicate profiles can be reused), look for a match now.
    ' NOTE: this IF/THEN block contains an Exit Function clause, and it will use it if a match is found.
    If ((m_NumOfCachedProfiles > 0) And matchDuplicates) Then
    
        Dim i As Long
        For i = 0 To m_NumOfCachedProfiles - 1
            If srcProfile.IsEqual(m_ProfileCache(i).fullProfile) Then
                AddProfileToCache = i
                Exit Function
            End If
        Next i
    
    End If
    
    'If we made it all the way here, assume we are good to add the current profile to our list
    With m_ProfileCache(m_NumOfCachedProfiles)
        Set .fullProfile = srcProfile
        .isSystemProfile = isSystemProfile
        .isPDDisplayProfile = isDisplayProfile
        .curDisplayID = associatedMonitorID
        .isWorkingSpaceProfile = isWorkingSpace
    End With
    
    AddProfileToCache = m_NumOfCachedProfiles
    m_NumOfCachedProfiles = m_NumOfCachedProfiles + 1

End Function

Public Function GetCachedProfile_ByIndex(ByVal profileindex As Long) As ICCProfileCache
    If (profileindex >= 0) And (profileindex < m_NumOfCachedProfiles) Then
        GetCachedProfile_ByIndex = m_ProfileCache(profileindex)
    End If
End Function

Public Function GetCachedDisplayProfileIndex_ByHandle(ByVal hMonitor As Long) As Long
    If (hMonitor <> 0) And (m_NumOfCachedProfiles <> 0) Then
        Dim i As Long
        For i = 0 To m_NumOfCachedProfiles - 1
            If m_ProfileCache(i).isPDDisplayProfile Then
                If (m_ProfileCache(i).curDisplayID = hMonitor) Then
                    GetCachedDisplayProfileIndex_ByHandle = i
                    Exit For
                End If
            End If
        Next i
    Else
        GetCachedDisplayProfileIndex_ByHandle = -1
    End If
End Function

'Given a unique profile string ID (as generated by GetUniqueProfileDescriptor_ByIndex(), below), try to find a matching profile.
' If no match is found, -1 is returned.
Public Function GetCachedProfileIndex_ByUniqueStringID(ByVal profileString As String) As Long
    
    GetCachedProfileIndex_ByUniqueStringID = -1
    
    If (m_NumOfCachedProfiles > 0) Then
        
        Dim i As Long, testStringID As String
        For i = 0 To m_NumOfCachedProfiles - 1
            
            testStringID = GetUniqueProfileDescriptor_ByIndex(i)
            
            If (StrComp(testStringID, profileString, vbBinaryCompare) = 0) Then
                GetCachedProfileIndex_ByUniqueStringID = i
                Exit For
            End If
            
        Next i
        
    End If
    
End Function

'If you want an immutable descriptor for a given profile, use this function.  It takes an index, and returns a (potentially lengthy)
' string that can be used to uniquely identify an ICC profile across sessions.
Public Function GetUniqueProfileDescriptor_ByIndex(ByVal profileindex As Long) As String
    
    If (profileindex >= 0) And (profileindex < m_NumOfCachedProfiles) Then
        With m_ProfileCache(profileindex)
            
            'If we've already calculed a unique identifier for this profile, reuse it
            If (Len(.profileStringID) <> 0) Then
                GetUniqueProfileDescriptor_ByIndex = .profileStringID
            
            'Unique IDs only need to be created once.  They are subsequently stored in .profileStringID.
            Else
            
                'Make sure an attached profile exists
                If (.lcmsProfileHandle = 0) Then
                    .lcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(.fullProfile.GetICCDataPointer, .fullProfile.GetICCDataSize)
                End If
                
                'Concatenate a bunch of descriptor strings, which forms a unique identifier
                GetUniqueProfileDescriptor_ByIndex = vbNullString
                
                Dim tmpString() As String
                ReDim tmpString(cmsInfoDescription To cmsInfoCopyright) As String
                
                Dim i As Long
                For i = cmsInfoDescription To cmsInfoCopyright
                    tmpString(i) = LittleCMS.LCMS_GetProfileInfoString(.lcmsProfileHandle, i)
                Next i
                
                For i = cmsInfoDescription To cmsInfoCopyright
                    .profileStringID = .profileStringID & "|-|" & tmpString(i)
                Next i
                
                GetUniqueProfileDescriptor_ByIndex = .profileStringID
                
            End If
            
        End With
    End If
    
End Function

'Release the full ICC cache.  Do not do this until PD is fully unloaded, or things will break.
Public Sub FreeProfileCache()
    
    'If one or more profiles exist in the cache, we must free any/all LittleCMS handles prior to exiting (or they will leak)
    If (m_NumOfCachedProfiles > 0) Then
    
        Dim i As Long
        For i = 0 To m_NumOfCachedProfiles - 1
            With m_ProfileCache(i)
                
                If (.thisWSToDisplayTransform24 <> 0) Then
                    LittleCMS.LCMS_DeleteTransform .thisWSToDisplayTransform24
                    .thisWSToDisplayTransform24 = 0
                End If
                
                If (.thisWSToDisplayTransform32 <> 0) Then
                    LittleCMS.LCMS_DeleteTransform .thisWSToDisplayTransform32
                    .thisWSToDisplayTransform32 = 0
                End If
                
                If (.lcmsProfileHandle <> 0) Then
                    LittleCMS.LCMS_CloseProfileHandle .lcmsProfileHandle
                    .lcmsProfileHandle = 0
                End If
                
                .indexOfDisplayTransform = -1
                
            End With
        Next i
    
    End If
    
    m_NumOfCachedProfiles = 0
    ReDim m_ProfileCache(0) As ICCProfileCache
    m_SystemProfileIndex = 0
    m_CurrentDisplayIndex = 0
    m_sRGBIndex = 0
    
End Sub

'Retrieve the current system color profile directory
Public Function GetSystemColorFolder() As String

    'Prepare a blank string to receive the profile path
    Dim bufferSize As Long
    bufferSize = MAX_PATH
    
    Dim tmpPathString As String
    tmpPathString = String$(bufferSize, 0&)
    
    'Use the GetColorDirectory function to request the location of the system color folder
    If GetColorDirectory(0&, StrPtr(tmpPathString), bufferSize) = 0 Then
        GetSystemColorFolder = ""
    Else
        Dim cUnicode As pdUnicode
        Set cUnicode = New pdUnicode
        GetSystemColorFolder = cUnicode.TrimNull(tmpPathString)
    End If

End Function

'Returns the path to the default color mangement profile file (ICC or WCS) currently in use by the system.
Public Function GetDefaultICCProfilePath() As String

    'Prepare a blank string to receive the profile path
    Dim filenameLength As Long
    filenameLength = MAX_PATH
    
    Dim tmpPathString As String
    tmpPathString = String$(MAX_PATH, 0&)
    
    'Using the desktop DC as our reference, request the filename of the currently in-use ICM profile (which should be the system default)
    If GetICMProfile(GetDC(0), filenameLength, StrPtr(tmpPathString)) = 0 Then
        GetDefaultICCProfilePath = ""
    Else
        Dim cUnicode As pdUnicode
        Set cUnicode = New pdUnicode
        GetDefaultICCProfilePath = cUnicode.TrimNull(tmpPathString)
    End If
    
End Function

'When the main PD window is moved, the window manager will trigger this function.  (Because the user can set color management
' on a per-monitor basis, we must keep track of which monitor contains this PD instance.)
Public Sub CheckParentMonitor(Optional ByVal suspendRedraw As Boolean = False, Optional ByVal forceRefresh As Boolean = False)
    
    Dim oldDisplayIndex As Long
    oldDisplayIndex = m_CurrentDisplayIndex
    
    'Use the API to determine the monitor with the largest intersect with this window
    Dim monitorCheck As Long
    monitorCheck = MonitorFromWindow(FormMain.hWnd, MONITOR_DEFAULTTONEAREST)
    
    'If the detected monitor does not match this one, update this window and refresh its image (if necessary)
    If (monitorCheck <> m_CurrentMonitor) Or forceRefresh Then
        
        m_CurrentMonitor = monitorCheck
        
        'Update the current display index to match
        If (m_DisplayCMMPolicy = DCM_SystemProfile) Then
            m_CurrentDisplayIndex = m_SystemProfileIndex
            
        'If the user has manually specified per-monitor display profiles, we want to cache each display profile in turn.
        ElseIf (m_DisplayCMMPolicy = DCM_CustomProfile) Then
            m_CurrentDisplayIndex = GetCachedDisplayProfileIndex_ByHandle(m_CurrentMonitor)
            If (m_CurrentDisplayIndex = -1) Then m_CurrentDisplayIndex = m_sRGBIndex
        Else
            m_CurrentDisplayIndex = -1
        End If
        
        'm_CurrentDisplayIndex now points at the relevant display profile in our ICC profile cache.
        ' If it has changed since the last refresh, we need to reset any existing "working space to display" transforms
        ' that do not match the current display.
        If ((m_CurrentDisplayIndex >= 0) And (m_CurrentDisplayIndex <> oldDisplayIndex)) Or forceRefresh Then
            
            Dim i As Long
            For i = 0 To m_NumOfCachedProfiles - 1
                
                'Only working space profiles need to be updated
                With m_ProfileCache(i)
                    If .isWorkingSpaceProfile Then
                        
                        'Check the current transform.  If it...
                        ' 1) does exist, and...
                        ' 2) it matches an old display index...
                        '... we need to erase it.  (A new transform will be created on-demand, as necessary.)
                        ' Note that the "forceRefresh" parameter also affects this; when TRUE, we always release existing transforms
                        If (.thisWSToDisplayTransform32 <> 0) Then
                            If (.indexOfDisplayTransform <> m_CurrentDisplayIndex) Or forceRefresh Then
                                LittleCMS.LCMS_DeleteTransform .thisWSToDisplayTransform32
                                .thisWSToDisplayTransform32 = 0
                                .indexOfDisplayTransform = -1
                            End If
                        End If
                        
                        If (.thisWSToDisplayTransform24 <> 0) Then
                            If (.indexOfDisplayTransform <> m_CurrentDisplayIndex) Or forceRefresh Then
                                LittleCMS.LCMS_DeleteTransform .thisWSToDisplayTransform24
                                .thisWSToDisplayTransform24 = 0
                                .indexOfDisplayTransform = -1
                            End If
                        End If
                        
                    End If
                End With
            
            Next i
            
            'As a convenience, note display changes in the debug log
            #If DEBUGMODE = 1 Then
                Dim tmpProfile As ICCProfileCache
                tmpProfile = GetCachedProfile_ByIndex(m_CurrentDisplayIndex)
                If (Not tmpProfile.fullProfile Is Nothing) Then pdDebug.LogAction "Monitor change detected, new profile is: " & tmpProfile.fullProfile.GetOriginalICCPath
            #End If
        
        End If
        
        'If the user doesn't want us to redraw anything to match the new profile, exit
        If suspendRedraw Then Exit Sub
        
        'Various on-screen elements are color-managed, so they need to be redrawn first.
        
        'Modern PD controls subclass color management changes, so all we need to do is post the matching message internally
        UserControl_Support.PostPDMessage WM_PD_COLOR_MANAGEMENT_CHANGE
        
        'If no images have been loaded, exit
        If (pdImages(g_CurrentImage) Is Nothing) Then Exit Sub
        
        'If an image has been loaded, and it is valid, redraw it now
        If (pdImages(g_CurrentImage).Width > 0) And (pdImages(g_CurrentImage).Height > 0) And (FormMain.WindowState <> vbMinimized) And (g_WindowManager.GetClientWidth(FormMain.hWnd) > 0) And pdImages(g_CurrentImage).IsActive Then
            Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        End If
        
    End If
    
End Sub

'Transform a given DIB from the specified working space (or sRGB, if no index is supplied) to the current display space.
' Do not call this if you don't know what you're doing, as it is *not* reversible.
Public Sub ApplyDisplayColorManagement(ByRef srcDIB As pdDIB, Optional ByVal srcWorkingSpaceIndex As Long = -1, Optional ByVal checkPremultiplication As Boolean = True)
    
    'Note that this function does nothing if the display is not currently color managed
    If (Not srcDIB Is Nothing) And (m_DisplayCMMPolicy <> DCM_NoManagement) Then
        
        ValidateWorkingSpaceDisplayTransform srcWorkingSpaceIndex, srcDIB
        If checkPremultiplication Then PreValidatePremultiplicationForSrcDIB srcDIB
        
        'Apply the transformation to the source DIB
        If (srcDIB.GetDIBColorDepth = 32) Then
            LittleCMS.LCMS_ApplyTransformToDIB srcDIB, m_ProfileCache(srcWorkingSpaceIndex).thisWSToDisplayTransform32
        Else
            LittleCMS.LCMS_ApplyTransformToDIB srcDIB, m_ProfileCache(srcWorkingSpaceIndex).thisWSToDisplayTransform24
        End If
        
        If checkPremultiplication Then PostValidatePremultiplicationForSrcDIB srcDIB
        
    End If
    
End Sub

'Transform some region of a given DIB from the specified working space (or sRGB, if no index is supplied) to the current display space.
' Do not call this if you don't know what you're doing, as it is *not* reversible.
Public Sub ApplyDisplayColorManagement_RectF(ByRef srcDIB As pdDIB, ByRef srcRectF As RECTF, Optional ByVal srcWorkingSpaceIndex As Long = -1, Optional ByVal checkPremultiplication As Boolean = True)
    
    'Note that this function does nothing if the display is not currently color managed
    If (Not srcDIB Is Nothing) And (m_DisplayCMMPolicy <> DCM_NoManagement) Then
        
        ValidateWorkingSpaceDisplayTransform srcWorkingSpaceIndex, srcDIB
        If checkPremultiplication Then PreValidatePremultiplicationForSrcDIB srcDIB
        
        'Apply the transformation to the source DIB
        If (srcDIB.GetDIBColorDepth = 32) Then
            LittleCMS.LCMS_ApplyTransformToDIB_RectF srcDIB, m_ProfileCache(srcWorkingSpaceIndex).thisWSToDisplayTransform32, srcRectF
        Else
            LittleCMS.LCMS_ApplyTransformToDIB_RectF srcDIB, m_ProfileCache(srcWorkingSpaceIndex).thisWSToDisplayTransform24, srcRectF
        End If
        
        If checkPremultiplication Then PostValidatePremultiplicationForSrcDIB srcDIB
        
    End If
    
End Sub

'Convert a single RGBA long from the specified working space (or sRGB, if no index is supplied) to the current display space.
' Note that if the source has been created via VB's RGB() function, bytes will be in the wrong order (compared to a display buffer).
' Notify this function accordingly, and it will handle the temporary transforms required.
Public Sub ApplyDisplayColorManagement_SingleColor(ByVal srcColor As Long, ByRef dstColor As Long, Optional ByVal srcWorkingSpaceIndex As Long = -1, Optional ByVal srcIsRGBLong As Boolean = True)
    
    'Start by mirroring the source color to the destination; this is our fallback result if something goes wrong
    ' (or if color management is entirely disabled.
    dstColor = srcColor
    
    'Note that this function does nothing if the display is not currently color managed
    If (m_DisplayCMMPolicy <> DCM_NoManagement) And g_IsProgramRunning Then
        
        ValidateWorkingSpaceDisplayTransform srcWorkingSpaceIndex, Nothing
        
        'Apply the transformation to the source color, with special handling if the source is a long created by VB's RGB() function
        If srcIsRGBLong Then
            
            Dim tmpRGBASrc As RGBQUAD, tmpRGBADst As RGBQUAD
            With tmpRGBASrc
                .alpha = 255
                .Red = Colors.ExtractRed(srcColor)
                .Green = Colors.ExtractGreen(srcColor)
                .Blue = Colors.ExtractBlue(srcColor)
            End With
            
            LittleCMS.LCMS_TransformArbitraryMemory VarPtr(tmpRGBASrc), VarPtr(tmpRGBADst), 1, m_ProfileCache(srcWorkingSpaceIndex).thisWSToDisplayTransform32
            
            With tmpRGBADst
                dstColor = RGB(.Red, .Green, .Blue)
            End With
            
        Else
            LittleCMS.LCMS_TransformArbitraryMemory VarPtr(srcColor), VarPtr(dstColor), 1, m_ProfileCache(srcWorkingSpaceIndex).thisWSToDisplayTransform32
        End If
        
    End If
    
End Sub

Private Sub PreValidatePremultiplicationForSrcDIB(ByRef srcDIB As pdDIB)
    
    'If the source DIB is premultiplied, it needs to be un-premultiplied first
    m_PreAlphaManagementRequired = False
    If (srcDIB.GetDIBColorDepth = 32) Then
        m_PreAlphaManagementRequired = srcDIB.GetAlphaPremultiplication
    End If
    
    If m_PreAlphaManagementRequired Then srcDIB.SetAlphaPremultiplication False
    
End Sub

Private Sub PostValidatePremultiplicationForSrcDIB(ByRef srcDIB As pdDIB)
    If m_PreAlphaManagementRequired Then srcDIB.SetAlphaPremultiplication True
    m_PreAlphaManagementRequired = False
End Sub

'Before applying a working space to display transform, call this function to validate all involved profiles and transforms
Private Sub ValidateWorkingSpaceDisplayTransform(ByRef srcWorkingSpaceIndex As Long, ByRef srcDIB As pdDIB)
    
    'If the caller doesn't specify a working space index, assume sRGB.  (NOTE: this should never happen, but better safe than sorry.)
    If (srcWorkingSpaceIndex < 0) Then srcWorkingSpaceIndex = m_sRGBIndex
    
    'Make sure a transform exists for the requested working space / display combination
    With m_ProfileCache(srcWorkingSpaceIndex)
        
        'Make sure an LCMS-compatible handle exists for the working space profile
        If (.lcmsProfileHandle = 0) Then
            .lcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(.fullProfile.GetICCDataPointer, .fullProfile.GetICCDataSize)
        End If
        
        'Make sure an LCMS-compatible handle exists for the display profile
        If (m_ProfileCache(m_CurrentDisplayIndex).lcmsProfileHandle = 0) Then
            m_ProfileCache(m_CurrentDisplayIndex).lcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(m_ProfileCache(m_CurrentDisplayIndex).fullProfile.GetICCDataPointer, m_ProfileCache(m_CurrentDisplayIndex).fullProfile.GetICCDataSize)
        End If
        
        'Make sure a valid transform exists for this bit-depth / working-space / display combination
        Dim use32bppPath As Boolean
        If (srcDIB Is Nothing) Then
            use32bppPath = True
        Else
            use32bppPath = CBool(srcDIB.GetDIBColorDepth = 32)
        End If
        
        If use32bppPath Then
        
            If (.thisWSToDisplayTransform32 = 0) Or (.indexOfDisplayTransform <> m_CurrentDisplayIndex) Then
                If (.thisWSToDisplayTransform32 <> 0) Then LittleCMS.LCMS_DeleteTransform .thisWSToDisplayTransform32
                .thisWSToDisplayTransform32 = LittleCMS.LCMS_CreateTwoProfileTransform(.lcmsProfileHandle, m_ProfileCache(m_CurrentDisplayIndex).lcmsProfileHandle, TYPE_BGRA_8, TYPE_BGRA_8, m_DisplayRenderIntent, cmsFLAGS_COPY_ALPHA)
                .indexOfDisplayTransform = m_CurrentDisplayIndex
            End If
        
        'Assume a 24-bpp source
        Else
            If (.thisWSToDisplayTransform24 = 0) Or (.indexOfDisplayTransform <> m_CurrentDisplayIndex) Then
                If (.thisWSToDisplayTransform24 <> 0) Then LittleCMS.LCMS_DeleteTransform .thisWSToDisplayTransform24
                .thisWSToDisplayTransform24 = LittleCMS.LCMS_CreateTwoProfileTransform(.lcmsProfileHandle, m_ProfileCache(m_CurrentDisplayIndex).lcmsProfileHandle, TYPE_BGR_8, TYPE_BGR_8, m_DisplayRenderIntent, 0&)
                .indexOfDisplayTransform = m_CurrentDisplayIndex
            End If
        End If
        
    End With

End Sub

''Shorthand way to activate color management for anything with a DC
'' NOTE: now that GDI is no longer used for color management, this function is obsolete
'Public Sub TurnOnDefaultColorManagement(ByVal targetDC As Long, ByVal targetHwnd As Long)
'
'    'Perform a quick check to see if the target DC is requesting sRGB management.  If it is, we can skip
'    ' color management entirely, because PD stores all RGB data in sRGB anyway.
'    'TODO: fix this
'    'If Not (g_UseSystemColorProfile And g_IsSystemColorProfileSRGB) Then
'    If (Not g_IsSystemColorProfileSRGB) Then
'        AssignDefaultColorProfileToObject targetHwnd, targetDC
'        TurnOnColorManagementForDC targetDC
'    End If
'
'End Sub

''Assign the default color profile (whether the system profile or the user profile) to any arbitrary object.  Note that the object
'' MUST have an hWnd and an hDC property for this to work.
'' NOTE: now that GDI is no longer used for color management, this function is obsolete
'Public Sub AssignDefaultColorProfileToObject(ByVal objectHWnd As Long, ByVal objectHDC As Long)
'
'    'If the current user setting is "use system color profile", our job is easy.
'    'TODO: fix this; g_UseSystemColorProfile is no longer used
'
'    'If g_UseSystemColorProfile Then
'    '    SetICMProfile objectHDC, StrPtr(m_currentSystemColorProfile)
'    'Else
'
'        'Use the form's containing monitor to retrieve a matching profile from the preferences file
'        If Len(m_CurrentDisplayProfile) <> 0 Then
'            SetICMProfile objectHDC, StrPtr(m_CurrentDisplayProfile)
'        Else
'            SetICMProfile objectHDC, StrPtr(m_currentSystemColorProfile)
'        End If
'
'    'End If
'
'    'If you would like to test this function on a standalone ICC profile (generally something bizarre, to help you know
'    ' that the function is working), use something similar to the code below.
'    'Dim TEST_ICM As String
'    'TEST_ICM = "C:\PhotoDemon v4\PhotoDemon\no_sync\Images from testers\jpegs\ICC\WhackedRGB.icc"
'    'SetICMProfile targetDC, StrPtr(TEST_ICM)
'
'End Sub
'
''Turn on color management for a specified device context
'' NOTE: now that GDI is no longer used for color management, this function is obsolete
'Public Sub TurnOnColorManagementForDC(ByVal dstDC As Long)
'    SetICMMode dstDC, ICM_ON
'End Sub
'
''Turn off color management for a specified device context
'' NOTE: now that GDI is no longer used for color management, this function is obsolete
'Public Sub TurnOffColorManagementForDC(ByVal dstDC As Long)
'    SetICMMode dstDC, ICM_OFF
'End Sub

'RGB to XYZ conversion using custom endpoints requires a special transform.  We cannot use Microsoft's built-in transform methods as they do
' not support variable white space endpoints (WTF, MICROSOFT).
'
'At present, this functionality is used for PNG files that specify their own cHRM (chromaticity) chunk.
'
'Note that this function supports transforms in *both* directions!  The optional treatEndpointsAsForwardValues can be set to TRUE to use the
' endpoint math when converting TO the XYZ space; if false, sRGB will be used for the RGB -> XYZ conversion, then the optional parameters
' will be used for the XYZ -> RGB conversion.  Directionality is important when working with filetypes (like PNG) that specify their own
' endpoints, as the endpoints define the reverse transform, not the forward one.  (Found this out the hard way, ugh.)
'
'Gamma is also optional; if none is specified, the default sRGB gamma transform value will be used.  Note that sRGB uses a two-part curve
' constructed around 2.4 - *not* a simple one-part 2.2 curve - so if you want 2.2 gamma, make sure you specify it!
'
'TODO: see if this can be migrated to LittleCMS instead; it will almost certainly be faster.
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
        PrepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
            
        'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
        Dim initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.GetDIBWidth - 1
        finalY = srcDIB.GetDIBHeight - 1
                
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim QuickVal As Long, qvDepth As Long
        qvDepth = srcDIB.GetDIBColorDepth \ 8
        
        'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
        ' based on the size of the area to be processed.
        Dim progBarCheck As Long
        If Not suppressMessages Then
            If modifyProgBarMax = -1 Then
                SetProgBarMax finalX
            Else
                SetProgBarMax modifyProgBarMax
            End If
            progBarCheck = FindBestProgBarValue()
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
                Colors.XYZtoRGB tmpX, tmpY, tmpZ, r, g, b
            
            'Reverse transform
            Else
            
                'Use sRGB for the initial XYZ conversion
                Colors.RGBtoXYZ r, g, b, tmpX, tmpY, tmpZ
            
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
                    If UserPressedESC() Then Exit For
                    SetProgBarVal x + modifyProgBarOffset
                End If
            End If
        Next x
                
        'With our work complete, point ImageData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
        Erase ImageData
        
        If g_cancelCurrentAction Then ConvertRGBUsingCustomEndpoints = False Else ConvertRGBUsingCustomEndpoints = True
        
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


'***

'IMPORTANT NOTE:
' All functions below this line are legacy Windows CMS lines.  They use various Windows functions to apply basic color management tasks.
' They are no longer used in PD except as last-resort fallbacks.  LittleCMS is used in their place.

'***


'Given a valid iccProfileArray (such as one stored in a pdICCProfile class), convert it to an internal Windows color profile
' handle, validate it, and return the result.  Returns a non-zero handle if successful.
Public Function LoadICCProfileFromMemory_WindowsCMS(ByVal profileArrayPointer As Long, ByVal profileArraySize As Long) As Long

    'Start by preparing an ICC_PROFILE header to use with the color management APIs
    Dim srcProfileHeader As ICC_PROFILE
    srcProfileHeader.dwType = PROFILE_MEMBUFFER
    srcProfileHeader.pProfileData = profileArrayPointer
    srcProfileHeader.cbDataSize = profileArraySize
    
    'Use that header to open a reference to an internal Windows color profile (which is required by all ICC-related API)
    LoadICCProfileFromMemory_WindowsCMS = OpenColorProfile(srcProfileHeader, PROFILE_READ, FILE_SHARE_READ, OPEN_EXISTING)
    
    If (LoadICCProfileFromMemory_WindowsCMS <> 0) Then
    
        'Validate the profile's XML as well; it is possible for a profile to be ill-formed, which means we cannot use it.
        Dim tmpCheck As Long
        If IsColorProfileValid(LoadICCProfileFromMemory_WindowsCMS, tmpCheck) = 0 Then
            Debug.Print "Color profile loaded succesfully, but XML failed to validate."
            CloseColorProfile LoadICCProfileFromMemory_WindowsCMS
            LoadICCProfileFromMemory_WindowsCMS = 0
        End If
        
    Else
        Debug.Print "ICC profile failed to load (OpenColorProfile failed with error #" & Err.LastDllError & ")."
    End If

End Function

'Given a valid ICC profile path, convert it to an internal Windows color profile handle, validate it,
' and return the result.  Returns a non-zero handle if successful.
Public Function LoadICCProfileFromFile_WindowsCMS(ByVal profilePath As String) As Long

    Dim cFile As pdFSO
    Set cFile = New pdFSO

    'Start by loading the specified path into a byte array
    Dim tmpProfileArray() As Byte
        
    If cFile.FileExist(profilePath) Then
        
        If Not cFile.LoadFileAsByteArray(profilePath, tmpProfileArray) Then
            LoadICCProfileFromFile_WindowsCMS = 0
            Exit Function
        End If
        
    Else
        LoadICCProfileFromFile_WindowsCMS = 0
        Exit Function
    End If

    'Next, prepare an ICC_PROFILE header to use with the color management APIs
    Dim srcProfileHeader As ICC_PROFILE
    srcProfileHeader.dwType = PROFILE_MEMBUFFER
    srcProfileHeader.pProfileData = VarPtr(tmpProfileArray(0))
    srcProfileHeader.cbDataSize = UBound(tmpProfileArray) + 1
    
    'Use that header to open a reference to an internal Windows color profile (which is required by all ICC-related API)
    LoadICCProfileFromFile_WindowsCMS = OpenColorProfile(srcProfileHeader, PROFILE_READ, FILE_SHARE_READ, OPEN_EXISTING)
    
    If (LoadICCProfileFromFile_WindowsCMS <> 0) Then
    
        'Validate the profile's XML as well; it is possible for a profile to be ill-formed, which means we cannot use it.
        Dim tmpCheck As Long
        If IsColorProfileValid(LoadICCProfileFromFile_WindowsCMS, tmpCheck) = 0 Then
            Debug.Print "Color profile loaded succesfully, but XML failed to validate."
            CloseColorProfile LoadICCProfileFromFile_WindowsCMS
            LoadICCProfileFromFile_WindowsCMS = 0
        End If
        
    Else
        Debug.Print "ICC profile failed to load (OpenColorProfile failed with error #" & Err.LastDllError & ")."
    End If

End Function

'Request a standard ICC profile from the OS.  Windows only provides two standard color profiles: sRGB (LCS_sRGB), and whatever
' the system default currently is (LCS_WINDOWS_COLOR_SPACE).  While probably not necessary, this function also validates the
' requested profile, just to be safe.
Public Function LoadStandardICCProfile_WindowsCMS(ByVal profileID As Long) As Long

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
    LoadStandardICCProfile_WindowsCMS = OpenColorProfile(dstProfileHeader, PROFILE_READ, FILE_SHARE_READ, OPEN_EXISTING)
    
    'It is highly unlikely (maybe even impossible?) for the system to return an invalid standard profile, but just to be
    ' safe, validate the XML.
    If (LoadStandardICCProfile_WindowsCMS <> 0) Then
    
        'Validate the profile's XML as well; it is possible for a profile to be ill-formed, which means we cannot use it.
        Dim tmpCheck As Long
        If IsColorProfileValid(LoadStandardICCProfile_WindowsCMS, tmpCheck) = 0 Then
            Debug.Print "Standard color profile loaded succesfully, but XML failed to validate."
            CloseColorProfile LoadStandardICCProfile_WindowsCMS
            LoadStandardICCProfile_WindowsCMS = 0
        End If
        
    Else
        Debug.Print "Standard ICC profile failed to load (GetStandardColorSpaceProfile failed with error #" & Err.LastDllError & ")."
    End If

End Function

'This function is just a thin wrapper to CloseColorProfile; however, using it allows us to keep various color-management
' DLLs nicely encapsulated within this module.
Public Sub ReleaseICCProfile_WindowsCMS(ByVal profileHandle As Long)
    CloseColorProfile profileHandle
End Sub

'Given a source profile, destination profile, and rendering intent, return a compatible transformation handle.
Public Function RequestProfileTransform_WindowsCMS(ByVal srcProfile As Long, ByVal dstProfile As Long, ByVal preferredIntent As RenderingIntents, Optional ByVal useEmbeddedIntent As Long = -1) As Long

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
    RequestProfileTransform_WindowsCMS = CreateMultiProfileTransform(ByVal VarPtr(profileMatrix(0)), 2&, ByVal VarPtr(intentMatrix(0)), 2&, (2 - g_ColorPerformance) + 1, INDEX_DONT_CARE)
    
    If (RequestProfileTransform_WindowsCMS = 0) Then
        Debug.Print "Requested color transformation could not be generated (Error #" & Err.LastDllError & ")."
    End If
    
End Function

'This function is just a thin wrapper to DeleteColorTransform; however, using it allows us to keep various color-management
' DLLs nicely encapsulated within this module.
Public Sub ReleaseColorTransform_WindowsCMS(ByVal transformHandle As Long)
    DeleteColorTransform transformHandle
End Sub

'Given a color transformation and a DIB, apply one to the other!  Returns TRUE if successful.
Public Function ApplyColorTransformToDIB_WindowsCMS(ByVal srcTransform As Long, ByRef dstDIB As pdDIB) As Boolean

    Dim transformCheck As Long
    
    With dstDIB
                
        'NOTE: note that I use BM_RGBTRIPLETS below, despite pdDIB DIBs most definitely being in BGR order.  This is an
        '       undocumented bug with Windows' color management engine!
        Dim bitDepthIdentifier As Long
        If .GetDIBColorDepth = 24 Then bitDepthIdentifier = BM_RGBTRIPLETS Else bitDepthIdentifier = BM_xRGBQUADS
                
        'TranslateBitmapBits handles the actual transformation for us.
        transformCheck = TranslateBitmapBits(srcTransform, .GetDIBPointer, bitDepthIdentifier, .GetDIBWidth, .GetDIBHeight, .GetDIBStride, .GetDIBPointer, bitDepthIdentifier, .GetDIBStride, ByVal 0&, 0&)
        
    End With
    
    If transformCheck = 0 Then
        ApplyColorTransformToDIB_WindowsCMS = False
        
        'Error #2021 is ERROR_COLORSPACE_MISMATCH: "The specified transform does not match the bitmap's color space."
        ' This is a known error when the source image was in CMYK format, because FreeImage (or GDI+) will have
        ' automatically converted the image to RGB at load-time.  Because the ICC profile is CMYK-specific, Windows will
        ' not be able to apply it to the image, as it is no longer in CMYK format!
        If (CLng(Err.LastDllError) = 2021) Then
            Debug.Print "Note: sRGB conversion already occurred."
        Else
            Debug.Print "ICC profile could not be applied.  Image remains in original profile. (Error #" & Err.LastDllError & ")."
        End If
        
    Else
        ApplyColorTransformToDIB_WindowsCMS = True
    End If

End Function

'Given a color transformation and two DIBs, fill one DIB with a transformed copy of the other!  Returns TRUE if successful.
Public Function ApplyColorTransformToTwoDIBs_WindowsCMS(ByVal srcTransform As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal srcFormat As Long, ByVal dstFormat As Long) As Boolean

    Dim transformCheck As Long
    
    'TranslateBitmapBits handles the actual transformation for us.
    transformCheck = TranslateBitmapBits(srcTransform, srcDIB.GetDIBPointer, srcFormat, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBStride, dstDIB.GetDIBPointer, dstFormat, dstDIB.GetDIBStride, ByVal 0&, 0&)
    
    If (transformCheck = 0) Then
        ApplyColorTransformToTwoDIBs_WindowsCMS = False
        
        'Error #2021 is ERROR_COLORSPACE_MISMATCH: "The specified transform does not match the bitmap's color space."
        ' This is a known error when the source image was in CMYK format, because FreeImage (or GDI+) will have
        ' automatically converted the image to RGB at load-time.  Because the ICC profile is CMYK-specific, Windows will
        ' not be able to apply it to the image, as it is no longer in CMYK format!
        If (CLng(Err.LastDllError) = 2021) Then
            Debug.Print "ICC profile could not be applied, because requested color spaces did not match supplied profile spaces."
        Else
            Debug.Print "ICC profile could not be applied.  Image remains in original profile. (Error #" & Err.LastDllError & ")."
        End If
        
    Else
        ApplyColorTransformToTwoDIBs_WindowsCMS = True
    End If

End Function

'Apply a CMYK transform between a 32bpp CMYK DIB and a 24bpp sRGB DIB.
Public Function ApplyCMYKTransform_WindowsCMS(ByVal iccProfilePointer As Long, ByVal iccProfileSize As Long, ByRef srcCMYKDIB As pdDIB, ByRef dstRGBDIB As pdDIB, Optional ByVal customSourceIntent As Long = -1) As Boolean

    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Using embedded ICC profile to convert image from CMYK to sRGB color space..."
    #End If
    
    'Use the ColorManagement module to convert the raw ICC profile into an internal Windows profile handle.  Note that
    ' this function will also validate the profile for us.
    Dim srcProfile As Long
    srcProfile = ColorManagement.LoadICCProfileFromMemory_WindowsCMS(iccProfilePointer, iccProfileSize)
    
    'If we successfully opened and validated our source profile, continue on to the next step!
    If (srcProfile <> 0) Then
    
        'Now it is time to determine our destination profile.  Because PhotoDemon operates on DIBs that default
        ' to the sRGB space, that's the profile we want to use for transformation.
            
        'Use the ColorManagement module to request a standard sRGB profile.
        Dim dstProfile As Long
        dstProfile = ColorManagement.LoadStandardICCProfile_WindowsCMS(LCS_sRGB)
        
        'It's highly unlikely that a request for a standard ICC profile will fail, but just be safe, double-check the
        ' returned handle before continuing.
        If (dstProfile <> 0) Then
            
            'We can now use our profile matrix to generate a transformation object, which we will use to directly modify
            ' the DIB's RGB values.
            Dim iccTransformation As Long
            iccTransformation = ColorManagement.RequestProfileTransform_WindowsCMS(srcProfile, dstProfile, INTENT_PERCEPTUAL, customSourceIntent)
            
            'If the transformation was generated successfully, carry on!
            If (iccTransformation <> 0) Then
                
                'The only transformation function relevant to PD involves the use of BitmapBits, so we will provide
                ' the API with direct access to our DIB bits.
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "CMYK to sRGB transform data created successfully.  Applying transform..."
                #End If
                
                'Note that a color format must be explicitly specified - we vary this contingent on the parent image's
                ' color depth.
                Dim transformCheck As Boolean
                transformCheck = ColorManagement.ApplyColorTransformToTwoDIBs_WindowsCMS(iccTransformation, srcCMYKDIB, dstRGBDIB, BM_KYMCQUADS, BM_RGBTRIPLETS)
                
                'If the transform was successful, pat ourselves on the back.
                If transformCheck Then
                    
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "CMYK to sRGB transformation successful."
                    #End If
                    
                    ApplyCMYKTransform_WindowsCMS = True
                    
                Else
                    Message "sRGB transform could not be applied.  Image remains in CMYK format."
                End If
                
                'Release our transformation
                ColorManagement.ReleaseColorTransform_WindowsCMS iccTransformation
                                
            Else
                Message "Both ICC profiles loaded successfully, but CMYK transformation could not be created."
                ApplyCMYKTransform_WindowsCMS = False
            End If
        
            ColorManagement.ReleaseICCProfile_WindowsCMS dstProfile
        
        Else
            Message "Could not obtain standard sRGB color profile.  CMYK transform abandoned."
            ApplyCMYKTransform_WindowsCMS = False
        End If
        
        ColorManagement.ReleaseICCProfile_WindowsCMS srcProfile
    
    Else
        Message "Embedded ICC profile is invalid.  CMYK transform could not be performed."
        ApplyCMYKTransform_WindowsCMS = False
    End If

End Function

'Compare two ICC profiles to determine equality.  Thank you to VB developer LaVolpe for this suggestion and original implementation.
Public Function AreColorProfilesEqual_WindowsCMS(ByVal profileHandle1 As Long, ByVal profileHandle2 As Long) As Boolean

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
    
    AreColorProfilesEqual_WindowsCMS = profilesEqual
    
End Function

'If an ICC profile is attached to a pdDIB object, apply said color profile to that pdDIB object.
Public Function ApplyICCtoPDDib_WindowsCMS(ByRef targetDIB As pdDIB) As Boolean
    
    ApplyICCtoPDDib_WindowsCMS = False
    
    'Before doing anything else, apply some failsafe checks
    If (targetDIB Is Nothing) Then Exit Function
    If (Not targetDIB.ICCProfile.HasICCData) Then Exit Function
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Using Windows ICM to convert pdDIB to current RGB working space..."
    #End If
    
    'Use the ColorManagement module to convert the raw ICC profile into an internal Windows profile handle.  Note that
    ' this function will also validate the profile for us.
    Dim srcProfile As Long
    srcProfile = ColorManagement.LoadICCProfileFromMemory_WindowsCMS(targetDIB.ICCProfile.GetICCDataPointer, targetDIB.ICCProfile.GetICCDataSize)
    
    'If we successfully opened and validated our source profile, continue on to the next step!
    If (srcProfile <> 0) Then
    
        'Now it is time to determine our destination profile.  Because PhotoDemon operates on DIBs that default
        ' to the sRGB space, that's the profile we want to use for transformation.
            
        'Use the ColorManagement module to request a standard sRGB profile.
        Dim dstProfile As Long
        dstProfile = ColorManagement.LoadStandardICCProfile_WindowsCMS(LCS_sRGB)
        
        'It's highly unlikely that a request for a standard ICC profile will fail, but just be safe, double-check the
        ' returned handle before continuing.
        If (dstProfile <> 0) Then
            
            'Before proceeding, check to see if the source and destination profiles are identical.  Some dSLRs will embed
            ' sRGB transforms in their JPEGs, and applying another sRGB transform atop them is a waste of time and resources.
            ' Thanks to VB developer LaVolpe for this suggestion.
            If (Not ColorManagement.AreColorProfilesEqual_WindowsCMS(srcProfile, dstProfile)) Then
            
                'We can now use our profile matrix to generate a transformation object, which we will use to directly modify
                ' the DIB's RGB values.
                Dim iccTransformation As Long
                iccTransformation = ColorManagement.RequestProfileTransform_WindowsCMS(srcProfile, dstProfile, INTENT_PERCEPTUAL, targetDIB.ICCProfile.GetSourceRenderIntent)
                
                'If the transformation was generated successfully, carry on!
                If (iccTransformation <> 0) Then
                    
                    'The only transformation function relevant to PD involves the use of BitmapBits, so we will provide
                    ' the API with direct access to our DIB bits.
                    
                    'Note that a color format must be explicitly specified - we vary this contingent on the parent image's
                    ' color depth.
                    Dim transformCheck As Boolean
                    transformCheck = ColorManagement.ApplyColorTransformToDIB_WindowsCMS(iccTransformation, targetDIB)
                    
                    'If the transform was successful, pat ourselves on the back.
                    If transformCheck Then
                    
                        #If DEBUGMODE = 1 Then
                            pdDebug.LogAction "ICC profile transformation successful.  Image is now sRGB."
                        #End If
                        
                        ApplyICCtoPDDib_WindowsCMS = True
                        targetDIB.ICCProfile.MarkSuccessfulProfileApplication
                        
                    Else
                        Message "ICC profile could not be applied.  Image remains in original profile."
                    End If
                    
                    'Release our transformation
                    ColorManagement.ReleaseColorTransform_WindowsCMS iccTransformation
                                    
                Else
                    Message "Both ICC profiles loaded successfully, but transformation could not be created."
                    ApplyICCtoPDDib_WindowsCMS = False
                End If
                
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "ICC transform is not required, because source and destination profiles are identical."
                #End If
                ApplyICCtoPDDib_WindowsCMS = True
            End If
        
            ColorManagement.ReleaseICCProfile_WindowsCMS dstProfile
        
        Else
            Message "Could not obtain standard sRGB color profile.  Color management has been disabled for this image."
            ApplyICCtoPDDib_WindowsCMS = False
        End If
        
        ColorManagement.ReleaseICCProfile_WindowsCMS srcProfile
    
    Else
        Message "Embedded ICC profile is invalid.  Color management has been disabled for this image."
        ApplyICCtoPDDib_WindowsCMS = False
    End If
    
End Function
