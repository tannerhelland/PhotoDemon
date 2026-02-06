Attribute VB_Name = "ColorManagement"
'***************************************************************************
'PhotoDemon ICC Profile Support Module
'Copyright 2013-2026 by Tanner Helland
'Created: 05/November/13
'Last updated: 24/April/18
'Last update: new helper function for exporting color profiles to file
'
'ICC profiles can be embedded in certain image file formats.  These profiles describe how to convert
' an image to a precisely defined reference space, while taking into account any pecularities of the
' device that captured the image (typically a camera).  From that reference space, we can then convert
' the image into any other device-specific color space (typically a monitor or printer).
'
'ICC profile handling is broken into three parts: extracting the profile from an image, using the
' extracted profile to convert an image into a reference "working space" (often sRGB, but HDR has
' different requirements), then activating color management for any user-facing DCs using the color
' profiles specified by the user.
'
'This class interacts heavily with pdICCProfile (a class for managing ICC profile data), and the
' LittleCMS plugin, which handles the bulk of PD's color management requirements.  (Past versions of
' this module used the built-in Windows color management module, known as "ICM" or "WCS" depending on
' the version, but I have since dropped all support for the Windows CMM as it's largely garbage.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit


'A handle (HMONITOR, specifically) to the main form's current monitor.  This value is updated by firing the
' CheckParentMonitor() function, below.
Private m_CurrentMonitor As Long

'APIs for retrieving system color profile folder(s)
Private Declare Function GetColorDirectory Lib "mscms" Alias "GetColorDirectoryW" (ByVal pMachineName As Long, ByVal ptrToBuffer As Long, ByRef pdwSize As Long) As Long
Private Declare Function MonitorFromWindow Lib "user32" (ByVal myHwnd As Long, ByVal dwFlags As Long) As Long

'Retrieves the filename of the color management file associated with a given DC
Private Declare Function GetICMProfile Lib "gdi32" Alias "GetICMProfileW" (ByVal hDC As Long, ByRef lpcbName As Long, ByVal ptrToBuffer As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Const MONITOR_DEFAULTTONEAREST As Long = &H2

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

'Current display render intent and BPC.  As of 7.0, the user can change these settings (although it doesn't make a
' lot of sense to do so, unless you're trying to do a poor-man's soft-proofing).
Private m_DisplayRenderIntent As LCMS_RENDERING_INTENT, m_DisplayBPC As Boolean

'When a pdImage object is created, we note the image's current color-management state.  Because the user can change
' the default PD color-management approach mid-session, different pdImage objects may be managed using different
' color management strategies.  The main user preference preference should only be referenced when an image is
' first loaded/created - after that point, use the preference embedded in the pdImage object instead.
Public Enum PD_ColorManagementState
    cms_NoManagement = 0
    cms_ProfileTagged = 1
    cms_ProfileConverted = 2
End Enum

#If False Then
    Private Const cms_NoManagement = 0, cms_ProfileTagged = 1, cms_ProfileConverted = 2
#End If

'If loaded successfully, the current system profile will be found at this index into the
' profile cache.  Note that this value *does not* support multiple monitors, due to flaws
' in the way Windows exposes the system profile (you can only easily retrieve the default
' monitor profile).  As such, if the user has custom-configured a monitor profile in the
' Tools > Options dialog, you shouldn't be using this value at all.
Private m_SystemProfileIndex As Long, m_SystemProfileHash As String

'Current display index.  This value is automatically refreshed by calls to CheckParentMonitor,
' below, and is used to support multimonitor systems.  On some configurations, it may be identical
' to m_SystemProfileIndex, above, or m_sRGBIndex, below.
Private m_CurrentDisplayIndex As Long

'sRGB profile index and hash value.  One valid sRGB profile is always loaded into memory, and it is used as a
' failsafe when things go horribly wrong (e.g. no system profile is configured, a requested display profile is
' corrupted or missing, etc).
Private m_sRGBIndex As Long, m_sRGBHash As String

'Cached profiles contain more information than a typical ICC profile, to make it easier to match profiles
' against each other.
Private Type ICCProfileCache
    
    'All profiles must be accompanied of a bytestream copy of their ICC profile contents.
    ' This allows us to do things like match duplicates, or look for duplicate source paths.
    FullProfile As pdICCProfile
    
    'Generally speaking, all profiles should also be accompanied by a LittleCMS handle to said profile.
    ' Note that these handles will leak if not manually released!
    LcmsProfileHandle As Long
    
    'Flags to help us shortcut searches for certain profile types
    IsSystemProfile As Boolean
    IsPDDisplayProfile As Boolean
    IsWorkingSpaceProfile As Boolean
    
    'This field is only used for display profiles; it is the HMONITOR corresponding to this display (used to match profiles
    ' to displays at run-time)
    CurDisplayID As Long
    
    'When used as part of the main viewport pipeline, working-space profiles cache a reference to a LittleCMS transform that
    ' translates between that working space and the current display profile.  This transform is optimized at creation time,
    ' and subsequently shared between all users of that working-space profile.  (The index of the matching display transform
    ' is also cached, so that we can detect future mismatches.)
    '
    'Note that 24-bpp and 32-bpp transforms are stored and optimized separately.  Both must be freed manually, if available.
    ThisWSToDisplayTransform24 As Long
    ThisWSToDisplayTransform32 As Long
    IndexOfDisplayTransform As Long
    
    'A unique string ID for this profile.  This is built by taking header information from the profile, and concatenating it
    ' together into something highly unique.  This string should not be presented to the user.  It *is* valid across sessions,
    ' and is used (for example) by pdLayer objects to uniquely identify their profiles when saved to file.
    ProfileStringID As String
    
    'A reusable hash that identifies this profile.  This string should not be presented to the user.  It *is* valid
    ' across sessions, and it can be used internally by PD to uniquely identify associated profiles.
    profileHash As String
    
End Type

'Current ICC Profile cache.  As of 7.0, profiles are cached here, in this sub, and any object that uses a profile
' simply receives an index into our cache.  This lets us reuse profiles for multiple objects.
Private Const INITIAL_PROFILE_CACHE_SIZE As Long = 8
Private m_NumOfCachedProfiles As Long
Private m_ProfileCache() As ICCProfileCache

'Some transforms require us to do pre- and post-conversion alpha management.  That management is tracked here, to prevent individual
' functions from needing to track it.
Private m_PreAlphaManagementRequired As Boolean

'To improve cache access, we generate unique hashes for each loaded profile.
Private m_Hasher As pdCrypto

'As of PD 2024.12, users can opt to disable both traditional color management, and format-specific color management
' (e.g. PNG gAMA/cHRM data).
Private m_UseEmbeddedICCProfiles As Boolean, m_UseEmbeddedLegacyProfiles   As Boolean

Public Sub UpdateColorManagementPreferences()
    m_UseEmbeddedICCProfiles = UserPrefs.GetPref_Boolean("ColorManagement", "allow-icc-profiles", True)
    m_UseEmbeddedLegacyProfiles = UserPrefs.GetPref_Boolean("ColorManagement", "allow-legacy-profiles", True)
End Sub

Public Function UseEmbeddedICCProfiles() As Boolean
    UseEmbeddedICCProfiles = m_UseEmbeddedICCProfiles
End Function

Public Function UseEmbeddedLegacyProfiles() As Boolean
    UseEmbeddedLegacyProfiles = m_UseEmbeddedLegacyProfiles
End Function

Public Function GetDisplayBPC() As Boolean
    GetDisplayBPC = UserPrefs.GetPref_Boolean("ColorManagement", "DisplayBPC", True)
End Function

Public Sub SetDisplayBPC(ByVal newValue As Boolean)
    UserPrefs.SetPref_Boolean "ColorManagement", "DisplayBPC", newValue
End Sub

Public Function GetDisplayColorManagementPreference() As DISPLAY_COLOR_MANAGEMENT
    GetDisplayColorManagementPreference = UserPrefs.GetPref_Long("ColorManagement", "DisplayCMMode", DCM_NoManagement)
    
    'Past PD versions used a true/false system to control this setting.  The old setting will be "-1" if the system
    ' color profile is in use.
    If (GetDisplayColorManagementPreference < 0) Then GetDisplayColorManagementPreference = DCM_SystemProfile
End Function

Public Sub SetDisplayColorManagementPreference(ByVal newPref As DISPLAY_COLOR_MANAGEMENT)
    UserPrefs.SetPref_Long "ColorManagement", "DisplayCMMode", newPref
End Sub

Public Function GetDisplayRenderingIntentPref() As LCMS_RENDERING_INTENT
    GetDisplayRenderingIntentPref = UserPrefs.GetPref_Long("ColorManagement", "DisplayRenderingIntent", INTENT_PERCEPTUAL)
End Function

Public Sub SetDisplayRenderingIntentPref(Optional ByVal newPref As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL)
    UserPrefs.SetPref_Long "ColorManagement", "DisplayRenderingIntent", newPref
End Sub

Public Function GetSRGBProfileHash() As String
    GetSRGBProfileHash = m_sRGBHash
End Function

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
            m_sRGBHash = AddProfileToCache(tmpProfile, False, False, False, False, True)
            m_sRGBIndex = GetProfileIndex_ByHash(m_sRGBHash)
            m_ProfileCache(m_sRGBIndex).LcmsProfileHandle = tmpHProfile
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
    m_DisplayCMMPolicy = ColorManagement.GetDisplayColorManagementPreference()
    m_DisplayRenderIntent = ColorManagement.GetDisplayRenderingIntentPref()
    m_DisplayBPC = ColorManagement.GetDisplayBPC()
    m_SystemProfileIndex = m_sRGBIndex
    
    'Start with the case where the user just wants us to use the default Windows system ICC profile.
    If (m_DisplayCMMPolicy = DCM_SystemProfile) Then
    
        m_currentSystemColorProfile = GetDefaultICCProfilePath()
        
        Set tmpProfile = New pdICCProfile
        If tmpProfile.LoadICCFromFile(m_currentSystemColorProfile) Then
            m_SystemProfileHash = AddProfileToCache(tmpProfile, False, True)
            m_SystemProfileIndex = GetProfileIndex_ByHash(m_SystemProfileHash)
            
            'Create an LCMS-compatible profile handle to match
            With m_ProfileCache(m_SystemProfileIndex)
                .LcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(.FullProfile.GetICCDataPointer, .FullProfile.GetICCDataSize)
            End With
            
        'If we fail to load the current system profile, there's really no good option for continuing.  The least
        ' of many evils is to simply point the current system profile at a default sRGB profile and hope for the best.
        Else
            PDDebug.LogAction "System ICC profile couldn't be loaded.  Reverting to default sRGB policy..."
            m_SystemProfileIndex = m_sRGBIndex
        End If
        
    'If the user has manually specified per-monitor display profiles, we want to cache each display profile in turn.
    ElseIf (m_DisplayCMMPolicy = DCM_CustomProfile) Then
        
        Dim tmpXML As pdXML
        Set tmpXML = New pdXML
        Dim i As Long, uniqueMonitorID As String, monICCPath As String
        Dim profileLoadedSuccessfully As Boolean, profileIndex As Long, profileHash As String
        
        For i = 0 To g_Displays.GetDisplayCount - 1
            
            profileLoadedSuccessfully = False
            
            With g_Displays.Displays(i)
                
                'Retrieve a unique ID for this display
                uniqueMonitorID = .GetUniqueDescriptor
                
                'Make it XML safe and look for a matching tag
                uniqueMonitorID = tmpXML.GetXMLSafeTagName(uniqueMonitorID)
                monICCPath = UserPrefs.GetPref_String("ColorManagement", "DisplayProfile_" & uniqueMonitorID, vbNullString)
                
                'If an ICC path exists for this display, attempt to load it
                If (LenB(monICCPath) <> 0) Then
                    
                    Set tmpProfile = New pdICCProfile
                    If tmpProfile.LoadICCFromFile(monICCPath) Then
                        
                        'Add the profile to our collection!
                        profileHash = AddProfileToCache(tmpProfile, False, False, True, .GetHandle)
                        profileIndex = GetProfileIndex_ByHash(profileHash)
                        
                        'Create an LCMS-compatible profile handle to match
                        If (profileIndex >= 0) Then
                            With m_ProfileCache(profileIndex)
                                .LcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(.FullProfile.GetICCDataPointer, .FullProfile.GetICCDataSize)
                            End With
                            profileLoadedSuccessfully = True
                        End If
                        
                    End If
                
                End If
                
                'If a profile was *not* loaded successfully, default to sRGB for this display.
                If (Not profileLoadedSuccessfully) And (m_sRGBIndex >= 0) Then
                    profileHash = AddProfileToCache(GetProfile_ByIndex(m_sRGBIndex).FullProfile, False, False, True, .GetHandle)
                    profileIndex = GetProfileIndex_ByHash(profileHash)
                    
                    If (profileIndex >= 0) Then
                        With m_ProfileCache(profileIndex)
                            .LcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(.FullProfile.GetICCDataPointer, .FullProfile.GetICCDataSize)
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

'Add a profile to the current cache.  The hash of said profile is returned; use that for any subsequent cache accesses.
Public Function AddProfileToCache(ByRef srcProfile As pdICCProfile, Optional ByVal matchDuplicates As Boolean = True, Optional ByVal IsSystemProfile As Boolean = False, Optional ByVal isDisplayProfile As Boolean = False, Optional ByVal associatedMonitorID As Long = 0, Optional ByVal isWorkingSpace As Boolean = False) As String
    
    'Make sure the cache exists and is large enough to hold another profile
    If (m_NumOfCachedProfiles = 0) Then
        ReDim m_ProfileCache(0 To INITIAL_PROFILE_CACHE_SIZE - 1) As ICCProfileCache
    Else
        If (m_NumOfCachedProfiles > UBound(m_ProfileCache)) Then ReDim m_ProfileCache(0 To (UBound(m_ProfileCache) * 2 + 1)) As ICCProfileCache
    End If
    
    'Profiles are quickly hashed; subsequent profile requests rely on this hash to return correct data
    If (m_Hasher Is Nothing) Then Set m_Hasher = New pdCrypto
    Dim profHash As String
    profHash = m_Hasher.QuickHash_AsString(srcProfile.GetICCDataPointer, srcProfile.GetICCDataSize, 16, PDCA_MD5)
    
    'Regardless of whether this profile already exists in our cache, we will return its hash value.  (This gives the
    ' caller a unique way to retrieve the profile in the future.)
    AddProfileToCache = profHash
    
    'If the user wants profile matching to occur (so that duplicate profiles can be reused), look for a match now.
    ' NOTE: this IF/THEN block contains an Exit Function clause, and it will use it if a match is found.
    If (matchDuplicates And (m_NumOfCachedProfiles > 0)) Then
        Dim i As Long
        For i = 0 To m_NumOfCachedProfiles - 1
            If (m_ProfileCache(i).profileHash = profHash) Then Exit Function
        Next i
    End If
    
    'If we made it all the way here, a match was *not* found.  This is a novel profile; add it to the cache.
    With m_ProfileCache(m_NumOfCachedProfiles)
        Set .FullProfile = srcProfile
        .profileHash = profHash
        .IsSystemProfile = IsSystemProfile
        .IsPDDisplayProfile = isDisplayProfile
        .CurDisplayID = associatedMonitorID
        .IsWorkingSpaceProfile = isWorkingSpace
    End With
    
    'Increment the profile cache size before exiting.  (Note that we already returned the
    m_NumOfCachedProfiles = m_NumOfCachedProfiles + 1
    
End Function

'Thin wrapper to AddProfileToCache(), above - but one that accepts an LCMS profile object.
Public Function AddLCMSProfileToCache(ByRef srcProfile As pdLCMSProfile, Optional ByVal matchDuplicates As Boolean = True, Optional ByVal IsSystemProfile As Boolean = False, Optional ByVal isDisplayProfile As Boolean = False, Optional ByVal associatedMonitorID As Long = 0, Optional ByVal isWorkingSpace As Boolean = False) As String
    Dim tmpProfile As pdICCProfile
    Set tmpProfile = New pdICCProfile
    If tmpProfile.LoadICCFromLCMSProfile(srcProfile) Then AddLCMSProfileToCache = AddProfileToCache(tmpProfile, matchDuplicates, IsSystemProfile, isDisplayProfile, associatedMonitorID, isWorkingSpace)
End Function

Public Function GetProfile_ByHash(ByRef srcHash As String) As pdICCProfile
    Dim i As Long
    For i = 0 To m_NumOfCachedProfiles - 1
        If (m_ProfileCache(i).profileHash = srcHash) Then
            Set GetProfile_ByHash = m_ProfileCache(i).FullProfile
            Exit Function
        End If
    Next i
End Function

Public Function GetProfile_ByIndex(ByVal profileIndex As Long) As ICCProfileCache
    If (profileIndex >= 0) And (profileIndex < m_NumOfCachedProfiles) Then
        GetProfile_ByIndex = m_ProfileCache(profileIndex)
    End If
End Function

Public Function GetProfile_Default() As pdICCProfile
    Dim i As Long
    i = ColorManagement.GetSRGBProfileIndex()
    If (i >= 0) Then Set GetProfile_Default = m_ProfileCache(i).FullProfile
End Function

Private Function GetProfileIndex_ByHash(ByRef srcHash As String) As Long
    Dim i As Long
    For i = 0 To m_NumOfCachedProfiles - 1
        If (m_ProfileCache(i).profileHash = srcHash) Then
            GetProfileIndex_ByHash = i
            Exit Function
        End If
    Next i
End Function

Public Function GetCachedDisplayProfileIndex_ByHandle(ByVal hMonitor As Long) As Long
    If (hMonitor <> 0) And (m_NumOfCachedProfiles <> 0) Then
        Dim i As Long
        For i = 0 To m_NumOfCachedProfiles - 1
            If m_ProfileCache(i).IsPDDisplayProfile Then
                If (m_ProfileCache(i).CurDisplayID = hMonitor) Then
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
Public Function GetCachedProfileIndex_ByUniqueStringID(ByRef profileString As String) As Long
    
    GetCachedProfileIndex_ByUniqueStringID = -1
    
    If (m_NumOfCachedProfiles > 0) Then
        
        Dim i As Long, testStringID As String
        For i = 0 To m_NumOfCachedProfiles - 1
            
            testStringID = GetUniqueProfileDescriptor_ByIndex(i)
            
            If Strings.StringsEqual(testStringID, profileString, False) Then
                GetCachedProfileIndex_ByUniqueStringID = i
                Exit For
            End If
            
        Next i
        
    End If
    
End Function

'If you want an immutable descriptor for a given profile, use this function.
' It takes an INDEX, and returns a (potentially lengthy) STRING that can be used to uniquely identify
' an ICC profile across sessions.
Public Function GetUniqueProfileDescriptor_ByIndex(ByVal profileIndex As Long) As String
    
    If (profileIndex >= 0) And (profileIndex < m_NumOfCachedProfiles) Then
        With m_ProfileCache(profileIndex)
            
            'If we've already calculed a unique identifier for this profile, reuse it
            If (LenB(.ProfileStringID) <> 0) Then
                GetUniqueProfileDescriptor_ByIndex = .ProfileStringID
            
            'Unique IDs only need to be created once.  They are subsequently stored in .ProfileStringID.
            Else
            
                'Make sure an attached profile exists
                If (.LcmsProfileHandle = 0) Then
                    .LcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(.FullProfile.GetICCDataPointer, .FullProfile.GetICCDataSize)
                End If
                
                'Concatenate a bunch of descriptor strings, which forms a unique identifier
                GetUniqueProfileDescriptor_ByIndex = vbNullString
                
                Dim tmpString() As String
                ReDim tmpString(cmsInfoDescription To cmsInfoCopyright) As String
                
                Dim i As Long
                For i = cmsInfoDescription To cmsInfoCopyright
                    tmpString(i) = LittleCMS.LCMS_GetProfileInfoString(.LcmsProfileHandle, i)
                Next i
                
                For i = cmsInfoDescription To cmsInfoCopyright
                    .ProfileStringID = .ProfileStringID & "|-|" & tmpString(i)
                Next i
                
                GetUniqueProfileDescriptor_ByIndex = .ProfileStringID
                
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
                
                If (.ThisWSToDisplayTransform24 <> 0) Then
                    LittleCMS.LCMS_DeleteTransform .ThisWSToDisplayTransform24
                    .ThisWSToDisplayTransform24 = 0
                End If
                
                If (.ThisWSToDisplayTransform32 <> 0) Then
                    LittleCMS.LCMS_DeleteTransform .ThisWSToDisplayTransform32
                    .ThisWSToDisplayTransform32 = 0
                End If
                
                If (.LcmsProfileHandle <> 0) Then
                    LittleCMS.LCMS_CloseProfileHandle .LcmsProfileHandle
                    .LcmsProfileHandle = 0
                End If
                
                .IndexOfDisplayTransform = -1
                
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
    If (GetColorDirectory(0&, StrPtr(tmpPathString), bufferSize) = 0) Then
        GetSystemColorFolder = vbNullString
    Else
        GetSystemColorFolder = Strings.TrimNull(tmpPathString)
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
    If (GetICMProfile(GetDC(0), filenameLength, StrPtr(tmpPathString)) <> 0) Then
        GetDefaultICCProfilePath = Strings.TrimNull(tmpPathString)
    Else
        GetDefaultICCProfilePath = vbNullString
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
                    If .IsWorkingSpaceProfile Then
                        
                        'Check the current transform.  If it...
                        ' 1) does exist, and...
                        ' 2) it matches an old display index...
                        '... we need to erase it.  (A new transform will be created on-demand, as necessary.)
                        ' Note that the "forceRefresh" parameter also affects this; when TRUE, we always release existing transforms
                        If (.ThisWSToDisplayTransform32 <> 0) Then
                            If (.IndexOfDisplayTransform <> m_CurrentDisplayIndex) Or forceRefresh Then
                                LittleCMS.LCMS_DeleteTransform .ThisWSToDisplayTransform32
                                .ThisWSToDisplayTransform32 = 0
                                .IndexOfDisplayTransform = -1
                            End If
                        End If
                        
                        If (.ThisWSToDisplayTransform24 <> 0) Then
                            If (.IndexOfDisplayTransform <> m_CurrentDisplayIndex) Or forceRefresh Then
                                LittleCMS.LCMS_DeleteTransform .ThisWSToDisplayTransform24
                                .ThisWSToDisplayTransform24 = 0
                                .IndexOfDisplayTransform = -1
                            End If
                        End If
                        
                    End If
                End With
            
            Next i
            
            'As a convenience, note display changes in the debug log
            If UserPrefs.GenerateDebugLogs Then
                Dim tmpProfile As ICCProfileCache
                tmpProfile = GetProfile_ByIndex(m_CurrentDisplayIndex)
                If (Not tmpProfile.FullProfile Is Nothing) Then PDDebug.LogAction "Monitor change detected, new profile is: " & tmpProfile.FullProfile.GetOriginalICCPath
            End If
        
        End If
        
        'If the user doesn't want us to redraw anything to match the new profile, exit
        If suspendRedraw Then Exit Sub
        
        'Various on-screen elements are color-managed, so they need to be redrawn first.
        
        'Modern PD controls subclass color management changes, so all we need to do is post the matching message internally
        UserControls.PostPDMessage WM_PD_COLOR_MANAGEMENT_CHANGE
        
        'If no images have been loaded, exit
        If (Not PDImages.IsImageActive()) Then Exit Sub
        
        'If an image has been loaded, and it is valid, redraw it now
        If (PDImages.GetActiveImage.Width > 0) And (PDImages.GetActiveImage.Height > 0) And (FormMain.WindowState <> vbMinimized) And (g_WindowManager.GetClientWidth(FormMain.hWnd) > 0) And PDImages.GetActiveImage.IsActive Then
            Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        End If
        
    End If
    
End Sub

'Apply an arbitrary profile to an arbitrary DIB.  You can either pass an explicit profile reference,
' or you can supply a hash to any valid profile inside PD's central profile cache.
Public Function ConvertDIBToSRGB(ByRef srcDIB As pdDIB, Optional ByRef srcProfile As pdICCProfile = Nothing, Optional ByRef useThisHashIDInstead As String = vbNullString) As Boolean
    
    If (Not PluginManager.IsPluginCurrentlyEnabled(CCP_LittleCMS)) Then
        PDDebug.LogAction "WARNING!  LittleCMS is missing, so color management has been disabled for this session."
        Exit Function
    End If
    
    'Make sure we have a valid source profile to work with
    If (srcProfile Is Nothing) And (LenB(useThisHashIDInstead) <> 0) Then Set srcProfile = ColorManagement.GetProfile_ByHash(useThisHashIDInstead)
    If (Not srcProfile Is Nothing) Then
        
        Dim srcLCMSProfile As pdLCMSProfile, dstLCMSProfile As pdLCMSProfile
        Set srcLCMSProfile = New pdLCMSProfile
        If srcLCMSProfile.CreateFromPDICCObject(srcProfile) Then
            
            Set dstLCMSProfile = New pdLCMSProfile
            dstLCMSProfile.CreateSRGBProfile
            
            Dim cTransform As pdLCMSTransform
            Set cTransform = New pdLCMSTransform
            If cTransform.CreateInPlaceTransformForDIB(srcDIB, srcLCMSProfile, dstLCMSProfile, INTENT_PERCEPTUAL, cmsFLAGS_COPY_ALPHA) Then
                ConvertDIBToSRGB = cTransform.ApplyTransformToPDDib(srcDIB)
            End If
            
        End If
        
    End If
    
End Function

'Transform a given DIB from the specified working space (or sRGB, if no index is supplied) to the current display space.
' Do not call this if you don't know what you're doing, as it is *not* reversible.
Public Sub ApplyDisplayColorManagement(ByRef srcDIB As pdDIB, Optional ByVal srcWorkingSpaceIndex As Long = -1, Optional ByVal checkPremultiplication As Boolean = True)
    
    'Note that this function does nothing if the display is not currently color managed
    If (Not srcDIB Is Nothing) And (m_DisplayCMMPolicy <> DCM_NoManagement) Then
        
        ValidateWorkingSpaceDisplayTransform srcWorkingSpaceIndex, srcDIB
        If checkPremultiplication Then PreValidatePremultiplicationForSrcDIB srcDIB
        
        'Apply the transformation to the source DIB
        If (srcDIB.GetDIBColorDepth = 32) Then
            LittleCMS.LCMS_ApplyTransformToDIB srcDIB, m_ProfileCache(srcWorkingSpaceIndex).ThisWSToDisplayTransform32
        Else
            LittleCMS.LCMS_ApplyTransformToDIB srcDIB, m_ProfileCache(srcWorkingSpaceIndex).ThisWSToDisplayTransform24
        End If
        
        If checkPremultiplication Then PostValidatePremultiplicationForSrcDIB srcDIB
        
    End If
    
End Sub

'Transform some region of a given DIB from the specified working space (or sRGB, if no index is supplied) to the current display space.
' Do not call this if you don't know what you're doing, as it is *not* reversible.
Public Sub ApplyDisplayColorManagement_RectF(ByRef srcDIB As pdDIB, ByRef srcRectF As RectF, Optional ByVal srcWorkingSpaceIndex As Long = -1, Optional ByVal checkPremultiplication As Boolean = True)
    
    'Note that this function does nothing if the display is not currently color managed
    If (Not srcDIB Is Nothing) And (m_DisplayCMMPolicy <> DCM_NoManagement) Then
        
        ValidateWorkingSpaceDisplayTransform srcWorkingSpaceIndex, srcDIB
        If checkPremultiplication Then PreValidatePremultiplicationForSrcDIB srcDIB
        
        'Apply the transformation to the source DIB
        If (srcDIB.GetDIBColorDepth = 32) Then
            LittleCMS.LCMS_ApplyTransformToDIB_RectF srcDIB, m_ProfileCache(srcWorkingSpaceIndex).ThisWSToDisplayTransform32, srcRectF
        Else
            LittleCMS.LCMS_ApplyTransformToDIB_RectF srcDIB, m_ProfileCache(srcWorkingSpaceIndex).ThisWSToDisplayTransform24, srcRectF
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
    If (m_DisplayCMMPolicy <> DCM_NoManagement) And PDMain.IsProgramRunning() Then
        
        ValidateWorkingSpaceDisplayTransform srcWorkingSpaceIndex, Nothing
        
        'Apply the transformation to the source color, with special handling if the source is a long created by VB's RGB() function
        If srcIsRGBLong Then
            
            Dim tmpRGBASrc As RGBQuad, tmpRGBADst As RGBQuad
            With tmpRGBASrc
                .Alpha = 255
                .Red = Colors.ExtractRed(srcColor)
                .Green = Colors.ExtractGreen(srcColor)
                .Blue = Colors.ExtractBlue(srcColor)
            End With
            
            LittleCMS.LCMS_TransformArbitraryMemory VarPtr(tmpRGBASrc), VarPtr(tmpRGBADst), 1, m_ProfileCache(srcWorkingSpaceIndex).ThisWSToDisplayTransform32
            
            With tmpRGBADst
                dstColor = RGB(.Red, .Green, .Blue)
            End With
            
        Else
            LittleCMS.LCMS_TransformArbitraryMemory VarPtr(srcColor), VarPtr(dstColor), 1, m_ProfileCache(srcWorkingSpaceIndex).ThisWSToDisplayTransform32
        End If
        
    End If
    
End Sub

Private Sub PreValidatePremultiplicationForSrcDIB(ByRef srcDIB As pdDIB)
    
    'If the source DIB is premultiplied, it needs to be un-premultiplied first
    m_PreAlphaManagementRequired = False
    If (srcDIB.GetDIBColorDepth = 32) Then m_PreAlphaManagementRequired = srcDIB.GetAlphaPremultiplication
    
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
        If (.LcmsProfileHandle = 0) Then
            .LcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(.FullProfile.GetICCDataPointer, .FullProfile.GetICCDataSize)
        End If
        
        'Make sure an LCMS-compatible handle exists for the display profile
        If (m_ProfileCache(m_CurrentDisplayIndex).LcmsProfileHandle = 0) Then
            m_ProfileCache(m_CurrentDisplayIndex).LcmsProfileHandle = LittleCMS.LCMS_LoadProfileFromMemory(m_ProfileCache(m_CurrentDisplayIndex).FullProfile.GetICCDataPointer, m_ProfileCache(m_CurrentDisplayIndex).FullProfile.GetICCDataSize)
        End If
        
        'Make sure a valid transform exists for this bit-depth / working-space / display combination
        Dim use32bppPath As Boolean
        If (srcDIB Is Nothing) Then
            use32bppPath = True
        Else
            use32bppPath = (srcDIB.GetDIBColorDepth = 32)
        End If
        
        Dim trnsFlags As LCMS_TRANSFORM_FLAGS
        If m_DisplayBPC Then trnsFlags = cmsFLAGS_BLACKPOINTCOMPENSATION
        
        'Verify the 32-bpp conversion handle
        If use32bppPath Then
        
            If (.ThisWSToDisplayTransform32 = 0) Or (.IndexOfDisplayTransform <> m_CurrentDisplayIndex) Then
                If (.ThisWSToDisplayTransform32 <> 0) Then LittleCMS.LCMS_DeleteTransform .ThisWSToDisplayTransform32
                .ThisWSToDisplayTransform32 = LittleCMS.LCMS_CreateTwoProfileTransform(.LcmsProfileHandle, m_ProfileCache(m_CurrentDisplayIndex).LcmsProfileHandle, TYPE_BGRA_8, TYPE_BGRA_8, m_DisplayRenderIntent, trnsFlags Or cmsFLAGS_COPY_ALPHA)
                .IndexOfDisplayTransform = m_CurrentDisplayIndex
            End If
        
        'Verify the 24-bpp conversion handle
        Else
            If (.ThisWSToDisplayTransform24 = 0) Or (.IndexOfDisplayTransform <> m_CurrentDisplayIndex) Then
                If (.ThisWSToDisplayTransform24 <> 0) Then LittleCMS.LCMS_DeleteTransform .ThisWSToDisplayTransform24
                .ThisWSToDisplayTransform24 = LittleCMS.LCMS_CreateTwoProfileTransform(.LcmsProfileHandle, m_ProfileCache(m_CurrentDisplayIndex).LcmsProfileHandle, TYPE_BGR_8, TYPE_BGR_8, m_DisplayRenderIntent, trnsFlags)
                .IndexOfDisplayTransform = m_CurrentDisplayIndex
            End If
        End If
        
    End With

End Sub

'Save a given pdImage's associated color profile to a standalone ICC file.
Public Function SaveImageProfileToFile(ByRef srcImage As pdImage) As Boolean
    
    'Failsafe checks
    If (srcImage Is Nothing) Then Exit Function
    If (LenB(srcImage.GetColorProfile_Original) = 0) Then Exit Function
    
    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    'Determine an initial folder.  This is easy - just grab the last "profile" path from the preferences file.
    Dim initialSaveFolder As String
    initialSaveFolder = UserPrefs.GetColorProfilePath()
    
    'Build a common dialog filter list
    Dim cdFilter As pdString, cdFilterExtensions As pdString
    Set cdFilter = New pdString
    Set cdFilterExtensions = New pdString
    
    cdFilter.Append g_Language.TranslateMessage("ICC profile") & " (.icc, .icm)|*.icc,*.icm"
    cdFilterExtensions.Append "icc"
    
    Dim cdIndex As Long
    cdIndex = 1
    
    'Suggest a file name.  At present, we just reuse the current image's name.
    Dim dstFilename As String
    dstFilename = srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    If (LenB(dstFilename) = 0) Then dstFilename = g_Language.TranslateMessage("New color profile")
    dstFilename = initialSaveFolder & dstFilename
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Export color profile")
    
    'Prep a common dialog interface
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    If saveDialog.GetSaveFileName(dstFilename, , True, cdFilter.ToString(), cdIndex, UserPrefs.GetColorProfilePath, cdTitle, cdFilterExtensions.ToString(), GetModalOwner().hWnd) Then
        
        'Update preferences
        UserPrefs.SetColorProfilePath Files.FileGetPath(dstFilename)
        
        'Pull the associated color profile into a byte array
        Dim srcProfile As pdICCProfile
        Set srcProfile = ColorManagement.GetProfile_ByHash(srcImage.GetColorProfile_Original)
        
        If (Not srcProfile Is Nothing) Then
            Dim profBytes() As Byte, profLen As Long
            srcProfile.GetProfileBytes profBytes, profLen
            If Files.FileCreateFromByteArray(profBytes, dstFilename, True) Then Message "Profile exported successfully."
        End If
        
    End If
    
    'Re-enable user input
    Interface.EnableUserInput
    
End Function
