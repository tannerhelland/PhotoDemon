VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Resource Manager
'Copyright 2016-2026 by Tanner Helland
'Created: 13/December/16
'Last updated: 21/June/22
'Last update: automatically extract 3D LUT (color lookup table) collection on first run
'
'PhotoDemon needs to include a whole swatch of custom resources.  These resources take up a lot of space,
' and we also need to pull different resources depending on things like screen DPI.  To simplify this process,
' we manage resources manually, rather than relying on an external .rc file to do it for us.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'If a resource file was loaded successfully, this will be set to TRUE.  You *must* check this value before
' attempting to retrieve individual resources.
Private m_ResourcesAvailable As Boolean

'Resources are stored in a persistent pdPackage.  You cannot rely on instantiation of this class for
' correct behavior; instead, query m_ResourcesAvailable.
Private m_Resources As pdPackageChunky

'Image resources are stored separately from other resources (because they have special compression needs).
' At run-time, image resource headers (XML) and pixel streams are pulled into two separate package managers.
' These only have to be decompressed once, at load-time; they are then kept uncompressed to maximize
' performance when retrieving individual resources.
Private m_ResImageHeaders As pdPackageChunky, m_ResImagePixels As pdPackageChunky

'We use a temporary DIB to pull resources from file, prior to applying any real-time coloring changes.
Private m_tmpDIB As pdDIB

'Some theme-related settings are cached to improve performance
Private m_ThemeIconColor As Long, m_ThemeIconColorMenu As Long, m_ThemeIconsMonochrome As Boolean

'Individual resources are always stored as two chunks: a raw data chunk, and an XML chunk that describes resource attributes.
' To improve performance, a dedicated XML parser is always created alongside this class.
Private m_XML As pdSerialize

'When loading image resources during startup, 99% of resources will be an identical size.  To cut down on the amount
' of memory we need to allocate, we try to reuse the same image buffer over and over again.
Private m_Palette() As RGBQuad, m_numOfPaletteEntries As Long
Private m_Pixels() As Byte, m_ResWidth As Long, m_ResHeight As Long

'Similarly, a persistent load buffer is also used when copying raw bytes out of individual resource blocks
Private m_tmpStream As pdStream

'Called early in PD's startup, this function checks to see if we've already extracted PD's resource file to a
' separate working folder.  If we have, grab the data from there instead of from the .exe.  (This is helpful
' during debugging.)
'
'In production builds, the initial collection is pulled from the .exe itself.
Friend Function LoadInitialResourceCollection() As Boolean
    
    Dim loadSuccess As Boolean
    
    'In developer builds, we preferentially load the file-based resource file from the App/PhotoDemon/Themes folder.
    ' (The resource file is updated frequently, and it's easier to work with a standalone file.)
    If (PD_BUILD_QUALITY <= PD_BETA) Then
    
        Dim srcResFile As String
        srcResFile = "Core_Icons.pdrc"
        loadSuccess = LoadResourceFile(UserPrefs.GetThemePath & srcResFile)
        
        'If the theme file doesn't exist, attempt to recover it.  (This may be necessary if the user used a dumb program
        ' like WinZip to unzip their PD download.)
        If (Not loadSuccess) Then
            AttemptResourceRecovery srcResFile
            loadSuccess = LoadResourceFile(UserPrefs.GetThemePath & srcResFile)
        End If
        
    End If
    
    'If the theme file was *not* loaded from a standalone file, load a good copy directly from this .exe's resource segment.
    If (Not loadSuccess) Then loadSuccess = LoadDefaultResource()
    
    'While we're here, cache the "monochrome icon" setting from the theme engine.
    If (Not g_Themer Is Nothing) Then m_ThemeIconsMonochrome = g_Themer.GetMonochromeIconSetting() Else m_ThemeIconsMonochrome = False
    
    LoadInitialResourceCollection = loadSuccess
    
End Function

Private Function AttemptResourceRecovery(ByVal resourceFile As String) As Boolean

    Dim resourceFilename As String
    resourceFilename = UserPrefs.GetThemePath(False) & resourceFile
    
    'pdFSO is used for all file interactions
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'See if the file exists.  If it does, great!  We can exit immediately.
    If Files.FileExists(resourceFilename) Then
        AttemptResourceRecovery = True
    
    'The file is missing.  Let's see if we can find it.
    Else
    
        RaiseResourceError "Resource file <" & resourceFile & "> is missing.  Scanning alternate folders...", True
        
        'See if the file exists in the base PD folder.  This can happen if a user unknowingly extracts
        ' the PD .zip without folders preserved.
        If Files.FileExists(UserPrefs.GetProgramPath & resourceFile) Then
            
            RaiseResourceError "Resource file found in the base PD folder.  Attempting to relocate...", True
            
            'Move the file to the proper folder
            If cFile.FileCopyW(UserPrefs.GetProgramPath & resourceFile, UserPrefs.GetThemePath & resourceFile) Then
                
                RaiseResourceError "Resource file relocated successfully.", True
                
                'Kill the old file instance
                Files.FileDeleteIfExists UserPrefs.GetProgramPath & resourceFile
                
                'Return success!
                AttemptResourceRecovery = True
            
            'The file couldn't be moved.  There's probably write issues with the folder structure, in which case
            ' this program session is pretty much doomed.  Exit now.
            Else
                RaiseResourceError "WARNING!  Resource file <" & resourceFile & "> could not be relocated."
                AttemptResourceRecovery = False
            End If
        
        'If the file doesn't exist in the base folder either, we're SOL.  Exit now.
        Else
            RaiseResourceError "Resource file <" & resourceFile & "> wasn't found in alternate locations.", True
            AttemptResourceRecovery = False
        End If
    
    End If
    
End Function

'Extract any default assets (e.g. standalone files stored inside the res segment) to their default location
' in the /Data subfolder.  Note that already extracted assets will *not* be re-extracted.
Friend Function ExtractDefaultAssets() As Boolean
    
    'Assets are stored as embedded pdPackageChunky instances.
    Dim cPackage As pdPackageChunky
    Set cPackage = New pdPackageChunky
    
    Dim dstFilename As String
    Dim tmpBytes() As Byte, tmpStream As pdStream, tmpChunkName As String, tmpChunkSize As Long
    
    'Extract the default gradient collection (but only if it's never been extracted before).
    If (Not UserPrefs.GetPref_Boolean("Assets", "ExtractedGradients", False)) Then
    
        PDDebug.LogAction "Gradient collection has never been extracted before.  Extracting now..."
        If Me.LoadGenericResource("gradients", tmpBytes) Then
            
            'Load the byte array into a temporary package manager
            If cPackage.OpenPackage_Memory(VarPtr(tmpBytes(0)), UBound(tmpBytes) + 1) Then
                
                'Gradient chunks are stored as name-value pairs
                Do While cPackage.GetNextChunk(tmpChunkName, tmpChunkSize, tmpStream)
                
                    'Ensure the chunk name is actually a "NAME" chunk
                    If (tmpChunkName = "NAME") Then
                        
                        'Convert the filename to a full path into the user's gradient folder
                        dstFilename = UserPrefs.GetGradientPath(True) & tmpStream.ReadString_UTF8(tmpChunkSize)
                        
                        'Extract the chunk's data
                        If cPackage.GetNextChunk(tmpChunkName, tmpChunkSize, tmpStream) Then
                            
                            'Ensure the chunk data is a "DATA" chunk
                            If (tmpChunkName = "DATA") Then
                                
                                'Write the chunk's contents to file
                                If (Not Files.FileCreateFromPtr(tmpStream.Peek_PointerOnly(0, tmpChunkSize), tmpChunkSize, dstFilename, True)) Then
                                    RaiseResourceError "ExtractDefaultAssets: failed to create target file " & dstFilename
                                End If
                                
                            End If
                            
                        Else
                            RaiseResourceError "bad gradient data: " & tmpChunkName
                        End If
                        
                    Else
                        RaiseResourceError "bad gradient name: " & tmpChunkName
                    End If
                
                'Iterate all remaining package items
                Loop
                
            Else
                RaiseResourceError "ExtractDefaultAssets: failed to extract gradient resource chunk"
            End If
            
        End If
        
        'Update the stored preference to ensure we do not extract these files again
        UserPrefs.SetPref_Boolean "Assets", "ExtractedGradients", True
        
    End If
    
    'Repeat the above steps for other assets
    
    'ICC profiles
    If (Not UserPrefs.GetPref_Boolean("Assets", "ExtractedColorProfiles", False)) Then
        PDDebug.LogAction "ICC profile collection has never been extracted before.  Extracting now..."
        If Me.LoadGenericResource("color_profiles", tmpBytes) Then
            If cPackage.OpenPackage_Memory(VarPtr(tmpBytes(0)), UBound(tmpBytes) + 1) Then
                Do While cPackage.GetNextChunk(tmpChunkName, tmpChunkSize, tmpStream)
                    If (tmpChunkName = "NAME") Then
                        dstFilename = UserPrefs.GetColorProfilePath() & tmpStream.ReadString_UTF8(tmpChunkSize)
                        If cPackage.GetNextChunk(tmpChunkName, tmpChunkSize, tmpStream) Then
                            If (tmpChunkName = "DATA") Then
                                If (Not Files.FileCreateFromPtr(tmpStream.Peek_PointerOnly(0, tmpChunkSize), tmpChunkSize, dstFilename, True)) Then
                                    RaiseResourceError "ExtractDefaultAssets: failed to create target file " & dstFilename
                                End If
                            End If
                        Else
                            RaiseResourceError "bad profile data: " & tmpChunkName
                        End If
                    Else
                        RaiseResourceError "bad gradient name: " & tmpChunkName
                    End If
                Loop
            Else
                RaiseResourceError "ExtractDefaultAssets: failed to extract profile resource chunk"
            End If
        End If
        UserPrefs.SetPref_Boolean "Assets", "ExtractedColorProfiles", True
    End If
    
    '3D LUTs
    If (Not UserPrefs.GetPref_Boolean("Assets", "Extracted3DLUTs", False)) Then
        PDDebug.LogAction "3D LUT collection has never been extracted before.  Extracting now..."
        If Me.LoadGenericResource("luts", tmpBytes) Then
            If cPackage.OpenPackage_Memory(VarPtr(tmpBytes(0)), UBound(tmpBytes) + 1) Then
                Do While cPackage.GetNextChunk(tmpChunkName, tmpChunkSize, tmpStream)
                    If (tmpChunkName = "NAME") Then
                        dstFilename = UserPrefs.GetLUTPath(True) & tmpStream.ReadString_UTF8(tmpChunkSize)
                        If cPackage.GetNextChunk(tmpChunkName, tmpChunkSize, tmpStream) Then
                            If (tmpChunkName = "DATA") Then
                                If (Not Files.FileCreateFromPtr(tmpStream.Peek_PointerOnly(0, tmpChunkSize), tmpChunkSize, dstFilename, True)) Then
                                    RaiseResourceError "ExtractDefaultAssets: failed to create target file " & dstFilename
                                End If
                            End If
                        Else
                            RaiseResourceError "bad lut data: " & tmpChunkName
                        End If
                    Else
                        RaiseResourceError "bad lut name: " & tmpChunkName
                    End If
                Loop
            Else
                RaiseResourceError "ExtractDefaultAssets: failed to extract profile resource chunk"
            End If
        End If
        UserPrefs.SetPref_Boolean "Assets", "Extracted3DLUTs", True
    End If
    
End Function

'Given a path to a .pdr file, attempt to load and validate it
Private Function LoadResourceFile(ByRef srcPath As String) As Boolean
        
    Set m_Resources = New pdPackageChunky
    
    If Files.FileExists(srcPath) Then
        
        RaiseResourceError "Loading core resource collection from file...", True
        
        'Load the file into memory (which allows us to overwrite the file, as necessary)
        Dim tmpBytes() As Byte
        Files.FileLoadAsByteArray srcPath, tmpBytes
        m_ResourcesAvailable = m_Resources.OpenPackage_Memory(VarPtr(tmpBytes(0)), UBound(tmpBytes) + 1, "PDRS", True)
        
    Else
        m_ResourcesAvailable = False
    End If
    
    LoadResourceFile = m_ResourcesAvailable
    
End Function

'Load the default resource collection directly from this .exe instance
Private Function LoadDefaultResource() As Boolean
        
    On Error GoTo NoResourcesAvailable
    
    RaiseResourceError "Falling back to internal resource collection...", True
        
    Set m_Resources = New pdPackageChunky
    
    'Pull the relevant resource out of memory
    Dim tmpRes() As Byte
    tmpRes = LoadResData("MAINRES", "CUSTOM")
    
    'Validate the resource
    m_ResourcesAvailable = m_Resources.OpenPackage_Memory(VarPtr(tmpRes(0)), UBound(tmpRes) + 1, "PDRS", True)
    LoadDefaultResource = m_ResourcesAvailable
    
    Exit Function
    
NoResourcesAvailable:
    RaiseResourceError "No resources found.  This session may not work as expected."
End Function

Friend Function AreResourcesAvailable() As Boolean
    AreResourcesAvailable = m_ResourcesAvailable
End Function

'When the user changes themes, the resource manager needs to be notified.  (Image resources are dynamically colored at run-time to
' match the current theme, so theme changes necessitate resource changes.)
Friend Sub NotifyThemeChange()
    m_ThemeIconsMonochrome = g_Themer.GetMonochromeIconSetting()
    m_ThemeIconColor = g_Themer.GetGenericUIColor(UI_IconMonochrome)
    m_ThemeIconColorMenu = g_Themer.GetGenericUIColor(UI_IconMonochromeMenu)
End Sub

'Load a generic resource (e.g. return it as a raw byte stream)
Friend Function LoadGenericResource(ByRef txtResName As String, ByRef dstBytes() As Byte) As Boolean
    
    LoadGenericResource = False
    
    If m_ResourcesAvailable Then
    
        LoadGenericResource = m_Resources.FindChunk_NameValuePair("NAME", txtResName, "DATA", m_tmpStream)
        
        If LoadGenericResource Then
            m_tmpStream.SetPosition 0, FILE_BEGIN
            m_tmpStream.ReadBytes dstBytes, m_tmpStream.GetStreamSize(), True
        Else
            RaiseResourceError "requested resource not found: " & txtResName
        End If
        
    End If

End Function

'Load an image-type resource.  Destination width and height must be manually specified.  If they are not specified, the imgae resource
' will be returned as-is.  Size is not consistent nor guaranteed to be correct.
'
'Optional padding and colors can also be specified, for places where icons are used in non-standard ways.  (Try to keep these to a minimum,
' as they are not guaranteed to work nicely with all themes.)
Friend Function LoadImageResource(ByRef imgResName As String, ByRef dstDIB As pdDIB, Optional ByVal desiredWidth As Long = 0, Optional ByVal desiredHeight As Long = 0, Optional ByVal desiredBorders As Single = 0!, Optional ByVal dstIsMenu As Boolean = False, Optional ByVal customColor As Long = -1, Optional ByVal suspendMonochrome As Boolean = False, Optional ByVal resampleAlgorithm As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal usePDResamplerInstead As PD_ResamplingFilter = rf_Automatic) As Boolean
    
    LoadImageResource = False
    
    If m_ResourcesAvailable Then
        
        'The first time an image resource is requested, load the two image packages from the central resource bank
        If (m_ResImageHeaders Is Nothing) Then
            If m_Resources.FindChunk_NameValuePair("NAME", "final_img_headers", "DATA", m_tmpStream) Then
                Set m_ResImageHeaders = New pdPackageChunky
                m_ResImageHeaders.OpenPackage_Memory m_tmpStream.Peek_PointerOnly(0, m_tmpStream.GetStreamSize()), m_tmpStream.GetStreamSize(), "IMGH", True
            Else
                PDDebug.LogAction "resource image header library missing"
            End If
        End If
        
        If (m_ResImagePixels Is Nothing) Then
            If m_Resources.FindChunk_NameValuePair("NAME", "final_img_pixels", "DATA", m_tmpStream) Then
                Set m_ResImagePixels = New pdPackageChunky
                m_ResImagePixels.OpenPackage_Memory m_tmpStream.Peek_PointerOnly(0, m_tmpStream.GetStreamSize()), m_tmpStream.GetStreamSize(), "IMGP", True
            Else
                PDDebug.LogAction "resource image pixel library missing"
            End If
        End If
        
        'See if this resource exists in the collection.
        If m_ResImageHeaders.FindChunk_NameValuePair("NAME", imgResName, "DATA", m_tmpStream) Then
            
            m_XML.SetParamString Strings.StringFromUTF8Ptr(m_tmpStream.Peek_PointerOnly(0, m_tmpStream.GetStreamSize()), m_tmpStream.GetStreamSize())
            
            'Retrieve the image's dimensions
            Dim imgWidth As Long, imgHeight As Long, imgBPP As Long
            imgWidth = m_XML.GetLong("w", 0, True)
            imgHeight = m_XML.GetLong("h", 0, True)
            imgBPP = m_XML.GetLong("bpp", 0, True)
            
            'Prep a temporary DIB (as we may need to resize the DIB to meet the user's request)
            If (m_tmpDIB Is Nothing) Then Set m_tmpDIB = New pdDIB
            
            Dim needToCreateDIB As Boolean: needToCreateDIB = False
            If (m_tmpDIB.GetDIBWidth <> imgWidth) Then
                needToCreateDIB = True
            ElseIf (m_tmpDIB.GetDIBHeight <> imgHeight) Then
                needToCreateDIB = True
            ElseIf (m_tmpDIB.GetDIBColorDepth <> imgBPP) Then
                needToCreateDIB = True
            End If
            
            If needToCreateDIB Then
                m_tmpDIB.CreateBlank imgWidth, imgHeight, imgBPP, 0, 0
                m_tmpDIB.SetInitialAlphaPremultiplicationState True
            End If
            
            'We now have to use one of two strategies to retrieve the DIB
            ' 1) Pull the actual DIB bits out of the file.  This is only an option if the image
            '    does *not* support run-time coloration.
            ' 2) Pull the DIB's alpha channel only out of the file.  If this image *does* support
            '    run-time coloration, there will not be any color data inside the file.
            If m_XML.GetBool("rt-clr", False, True) Then
                
                'Reuse the same module-level pixel array, to try and reduce unnecessary allocations
                If (m_ResWidth <> imgWidth) Or (m_ResHeight <> imgHeight) Then
                    m_ResWidth = imgWidth
                    m_ResHeight = imgHeight
                    ReDim m_Pixels(0 To imgWidth - 1, 0 To imgHeight - 1) As Byte
                End If
                
                LoadImageResource = m_ResImagePixels.FindChunk_NameValuePair("NAME", imgResName, "DATA", Nothing, VarPtr(m_Pixels(0, 0)))
                If LoadImageResource Then LoadImageResource = DIBs.ApplyTransparencyTable(m_tmpDIB, m_Pixels)
                
            Else
                
                'See if the DIB was stored using a palette system
                If m_XML.GetBool("uses-palette", False, True) Then
                
                    'Retrieve the palette count, and prep a palette and 8-bpp array
                    Dim numColors As Long
                    numColors = m_XML.GetLong("palette-size", 0, True)
                    
                    'Failsafe check only
                    If (numColors > 0) Then
                    
                        'Reuse the same module-level arrays, to try and reduce unnecessary allocations
                        If (m_numOfPaletteEntries < numColors) Then
                            m_numOfPaletteEntries = numColors
                            ReDim m_Palette(0 To numColors - 1) As RGBQuad
                        End If
                        
                        If (m_ResWidth <> imgWidth) Or (m_ResHeight <> imgHeight) Then
                            m_ResWidth = imgWidth
                            m_ResHeight = imgHeight
                            ReDim m_Pixels(0 To imgWidth - 1, 0 To imgHeight - 1) As Byte
                        End If
                        
                        If m_ResImagePixels.FindChunk_NameValuePair("NAME", imgResName, "DATA", m_tmpStream) Then
                            
                            'Copy the palette and pixel data into place
                            m_tmpStream.SetPosition 0, FILE_BEGIN
                            m_tmpStream.ReadBytesToBarePointer VarPtr(m_Palette(0)), numColors * 4
                            m_tmpStream.ReadBytesToBarePointer VarPtr(m_Pixels(0, 0)), imgWidth * imgHeight
                               
                            'Build a matching 32-bpp DIB from the palette and pixel data
                            LoadImageResource = DIBs.GetRGBADIB_FromPalette(m_tmpDIB, numColors, m_Palette, m_Pixels)
                            
                        End If
                    
                    End If
                
                Else
                    Dim tmpDIBPointer As Long, tmpDIBLength As Long
                    m_tmpDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                    LoadImageResource = m_ResImagePixels.FindChunk_NameValuePair("NAME", imgResName, "DATA", Nothing, tmpDIBPointer)
                End If
                
            End If
        
            'At present, all resources contain premultiplied alpha, so force the corresponding state now
            m_tmpDIB.SetInitialAlphaPremultiplicationState True
            
            'Resize the DIB into the destination
            If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
            If (desiredWidth = 0) Then desiredWidth = 16
            If (desiredHeight = 0) Then desiredHeight = 16
            
            If (dstDIB.GetDIBWidth <> desiredWidth) Or (dstDIB.GetDIBHeight <> desiredHeight) Then
                dstDIB.CreateBlank desiredWidth, desiredHeight, 32, 0, 0
            Else
                dstDIB.ResetDIB 0
            End If
            
            dstDIB.SetInitialAlphaPremultiplicationState True
            
            If (usePDResamplerInstead = rf_Automatic) Then
                GDI_Plus.GDIPlus_StretchBlt dstDIB, desiredBorders, desiredBorders, desiredWidth - desiredBorders * 2, desiredHeight - desiredBorders * 2, m_tmpDIB, 0, 0, imgWidth, imgHeight, , resampleAlgorithm, , True, , True
            
            'In a few special places, PD preferentially uses its own internal resampling engine.
            ' Note that - critically - this feature does *not* support the input border parameters.
            Else
                Resampling.ResampleImageI dstDIB, m_tmpDIB, desiredWidth, desiredHeight, usePDResamplerInstead, False
            End If
            
            'Now, we use a hierarchy of settings to determine how to color this particular icon.
            
            'First: if the user wants monochrome icons, this overrides all other color settings.
            Dim targetColor As Long
            
            If m_ThemeIconsMonochrome And (Not suspendMonochrome) Then
                If dstIsMenu Then targetColor = m_ThemeIconColorMenu Else targetColor = m_ThemeIconColor
                DIBs.ColorizeDIB dstDIB, targetColor
            
            'If the user does *not* want monochrome icons, we have more flexibility in how we deal with coloration.
            Else
            
                'If a custom color was specified, apply it now.
                If (customColor <> -1) Then
                    DIBs.ColorizeDIB dstDIB, customColor
                Else
                
                    'If the image supports real-time coloration, apply it now (based on the currently selected theme).
                    If m_XML.GetBool("rt-clr", False, True) Then
                    
                        'Retrieve the image's color.  (NOTE: the specified color will be overridden with
                        ' monochrome if the monochrome icon preference is active; see the branch above.)
                        If dstIsMenu Then
                            If m_XML.GetBool("rt-clrmenu", False, True) Then
                                targetColor = m_XML.GetLong("clr-m", 0, True)
                            Else
                                targetColor = m_XML.GetLong("clr-l", 0, True)
                            End If
                        Else
                            If ((g_Themer.GetCurrentThemeClass = PDTC_Light) Or dstIsMenu) Then
                                targetColor = m_XML.GetLong("clr-l", 0, True)
                            ElseIf (g_Themer.GetCurrentThemeClass = PDTC_Dark) Then
                                targetColor = m_XML.GetLong("clr-d", 0, True)
                            End If
                        End If
                    
                        DIBs.ColorizeDIB dstDIB, targetColor
                    
                    End If
                    
                End If
                
            End If
        
            'Before returning, free the target DIB from its DC.  (There's a chance our caller won't
            ' use the DIB right away, and this keeps an unnecessary GDI object from being created.)
            dstDIB.FreeFromDC
            
        'resource doesn't exist
        Else
            RaiseResourceError "Resource header missing: " & imgResName
        End If
    
    'central resource file isn't available
    Else
        RaiseResourceError "Resources aren't available (executable corrupt?)"
    End If
    
End Function

Friend Function LoadTextResource(ByRef txtResName As String, ByRef dstString As String) As Boolean
    
    LoadTextResource = False
    
    If m_ResourcesAvailable Then
        LoadTextResource = m_Resources.FindChunk_NameValuePair("NAME", txtResName, "DATA", m_tmpStream)
        If LoadTextResource Then
            m_tmpStream.SetPosition 0, FILE_BEGIN
            dstString = m_tmpStream.ReadString_UTF8(m_tmpStream.GetStreamSize())
        Else
            RaiseResourceError "couldn't find text resource: " & txtResName
        End If
    End If

End Function

Private Sub RaiseResourceError(ByVal msgError As String, Optional ByVal msgIsNonErrorFeedback As Boolean = False)
    If msgIsNonErrorFeedback Then
        PDDebug.LogAction "pdResources reported: " & msgError
    Else
        PDDebug.LogAction "WARNING!  pdResources error: " & msgError
    End If
End Sub

Private Sub Class_Initialize()
    
    m_ResourcesAvailable = False
    m_ThemeIconsMonochrome = False
    
    Set m_XML = New pdSerialize
    
End Sub
