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
'Copyright 2016-2017 by Tanner Helland
'Created: 13/December/16
'Last updated: 11/August/17
'Last update: use class-level caches for loading resources, which should reduce memory churn
'
'PhotoDemon needs to include a whole swatch of custom resources.  These resources take up a lot of space,
' and we also need to pull different resources depending on things like screen DPI.  To simplify this process,
' we manage resources manually, rather than relying on an external .rc file to do it for us.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'If a resource file was loaded successfully, this will be set to TRUE.  You *must* check this value before
' attempting to retrieve individual resources.
Private m_ResourcesAvailable As Boolean

'Resources are stored in a persistent pdPackager instance.  You cannot rely on instantiation of this class
' for correct behavior; instead, query m_ResourcesAvailable.
Private m_Resources As pdPackager

'We use a temporary DIB to pull resources from file, prior to applying any real-time coloring changes.
Private m_tmpDIB As pdDIB

'Some theme-related settings are cached to improve performance
Private m_ThemeIconColor As Long, m_ThemeIconColorMenu As Long, m_ThemeIconsMonochrome As Boolean

'Individual resources are always stored as two chunks: a raw data chunk, and an XML chunk that describes resource attributes.
' To improve performance, a dedicated XML parser is always created alongside this class.
Private m_XML As pdXML

'When loading image resources during startup, 99% of resources will be an identical size.  To cut down on the amount
' of memory we need to allocate, we try to reuse the same image buffer over and over again.
Private m_Palette() As RGBQUAD, m_numOfPaletteEntries As Long
Private m_Pixels() As Byte, m_ResWidth As Long, m_ResHeight As Long

'Similarly, a persistent load buffer is also used when copying raw bytes out of individual resource blocks
Private m_Buffer() As Byte

'Called early in PD's startup, this function checks to see if we've already extracted PD's resource file to a
' separate working folder.  If we have, grab the data from there instead of from the .exe.  (This is helpful
' during debugging.)
'
'In production builds, the initial collection is pulled from the .exe itself.
Friend Function LoadInitialResourceCollection() As Boolean
    
    Dim loadSuccess As Boolean
    
    'In developer builds, we preferentially load the file-based resource file from the App/PhotoDemon/Themes folder.
    ' (The resource file is updated frequently, and it's easier to work with a standalone file.)
    Dim srcResFile As String
    srcResFile = "Core_Icons.pdrc"
    loadSuccess = LoadResourceFile(g_UserPreferences.GetThemePath & srcResFile)
        
    'If the theme file doesn't exist, attempt to recover it.  (This may be necessary if the user used a dumb program
    ' like WinZip to unzip their PD download.)
    If (Not loadSuccess) Then
        AttemptResourceRecovery srcResFile
        loadSuccess = LoadResourceFile(g_UserPreferences.GetThemePath & srcResFile)
    End If
    
    'If the theme file was *not* loaded from a standalone file, load a good copy directly from this .exe's resource segment.
    If (Not loadSuccess) Then loadSuccess = LoadDefaultResource()
    
    'While we're here, cache the "monochrome icon" setting from the theme engine.
    If (Not g_Themer Is Nothing) Then m_ThemeIconsMonochrome = g_Themer.GetMonochromeIconSetting() Else m_ThemeIconsMonochrome = False
    
    LoadInitialResourceCollection = loadSuccess
    
End Function

Private Function AttemptResourceRecovery(ByVal resourceFile As String) As Boolean

    Dim resourceFilename As String
    resourceFilename = g_UserPreferences.GetThemePath(False) & resourceFile
    
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
        If Files.FileExists(g_UserPreferences.GetProgramPath & resourceFile) Then
            
            RaiseResourceError "Resource file found in the base PD folder.  Attempting to relocate...", True
            
            'Move the file to the proper folder
            If cFile.FileCopyW(g_UserPreferences.GetProgramPath & resourceFile, g_UserPreferences.GetThemePath & resourceFile) Then
                
                RaiseResourceError "Resource file relocated successfully.", True
                
                'Kill the old file instance
                Files.FileDeleteIfExists g_UserPreferences.GetProgramPath & resourceFile
                
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

'Given a path to a .pdr file, attempt to load and validate it
Private Function LoadResourceFile(ByRef srcPath As String) As Boolean
        
    Set m_Resources = New pdPackager
    
    If Files.FileExists(srcPath) Then
        RaiseResourceError "Loading core resource collection from file...", True
        m_ResourcesAvailable = m_Resources.ReadPackageFromFile(srcPath, PD_RES_IDENTIFIER, PD_SM_MemoryBacked)
    Else
        m_ResourcesAvailable = False
    End If
    
    LoadResourceFile = m_ResourcesAvailable
    
End Function

'Load the default resource collection directly from this .exe instance
Private Function LoadDefaultResource() As Boolean
        
    On Error GoTo NoResourcesAvailable
    
    RaiseResourceError "Falling back to internal resource collection...", True
        
    Set m_Resources = New pdPackager
    
    'Pull the relevant resource out of memory
    Dim tmpRes() As Byte
    tmpRes = LoadResData("MAINRES", "CUSTOM")
    
    'Validate the resource
    m_ResourcesAvailable = m_Resources.ReadPackageFromMemory(VarPtr(tmpRes(0)), UBound(tmpRes) + 1, PD_RES_IDENTIFIER)
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

'Load an image-type resource.  Destination width and height must be manually specified.  If they are not specified, the imgae resource
' will be returned as-is.  Size is not consistent nor guaranteed to be correct.
'
'Optional padding and colors can also be specified, for places where icons are used in non-standard ways.  (Try to keep these to a minimum,
' as they are not guaranteed to work nicely with all themes.)
Friend Function LoadImageResource(ByRef imgResName As String, ByRef dstDIB As pdDIB, Optional ByVal desiredWidth As Long = 0, Optional ByVal desiredHeight As Long = 0, Optional ByVal desiredBorders As Single = 0#, Optional ByVal dstIsMenu As Boolean = False, Optional ByVal customColor As Long = -1, Optional ByVal suspendMonochrome As Boolean = False) As Boolean
    
    LoadImageResource = False
    
    If m_ResourcesAvailable Then
    
        'See if this resource exists in the collection.  (It may not, as I'm still in the process of migrating
        ' all image resources to the new resource format.)
        Dim nodeIndex As Long
        nodeIndex = m_Resources.GetNodeIndexFromName(imgResName, vbBinaryCompare)
        
        '-1 is returned if the node doesn't exist in the collection
        If (nodeIndex >= 0) Then
            
            'Retrieve the header XML packet for this node
            Dim retString As String, retSize As Long
            If m_Resources.GetNodeDataByIndex(nodeIndex, True, m_Buffer, False, retSize) Then
                
                'Copy the received bytes into a string, then initialize an XML object from it
                retString = Space$(retSize \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(m_Buffer(0)), retSize
                
                If m_XML.LoadXMLFromString(retString) Then
                
                    'Retrieve the image's dimensions
                    Dim imgWidth As Long, imgHeight As Long, imgBPP As Long
                    imgWidth = m_XML.GetUniqueTag_Long("w")
                    imgHeight = m_XML.GetUniqueTag_Long("h")
                    imgBPP = m_XML.GetUniqueTag_Long("bpp")
                    
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
                    If m_XML.GetUniqueTag_Boolean("rt-clr") Then
                        
                        'Reuse the same module-level pixel array, to try and reduce unnecessary allocations
                        If (m_ResWidth <> imgWidth) Or (m_ResHeight <> imgHeight) Then
                            m_ResWidth = imgWidth
                            m_ResHeight = imgHeight
                            ReDim m_Pixels(0 To imgWidth - 1, 0 To imgHeight - 1) As Byte
                        End If
                        
                        LoadImageResource = m_Resources.GetNodeDataByIndex_UnsafeDstPointer(nodeIndex, False, VarPtr(m_Pixels(0, 0)))
                        If LoadImageResource Then LoadImageResource = DIBs.ApplyTransparencyTable(m_tmpDIB, m_Pixels)
                        
                    Else
                        
                        'See if the DIB was stored using a palette system
                        If m_XML.GetUniqueTag_Boolean("uses-palette", False) Then
                        
                            'Retrieve the palette count, and prep a palette and 8-bpp array
                            Dim numColors As Long
                            numColors = m_XML.GetUniqueTag_Long("palette-size")
                            
                            If (numColors > 0) Then
                                
                                'Reuse the same module-level arrays, to try and reduce unnecessary allocations
                                If (m_numOfPaletteEntries < numColors) Then
                                    m_numOfPaletteEntries = numColors
                                    ReDim m_Palette(0 To numColors - 1) As RGBQUAD
                                End If
                                
                                If (m_ResWidth <> imgWidth) Or (m_ResHeight <> imgHeight) Then
                                    m_ResWidth = imgWidth
                                    m_ResHeight = imgHeight
                                    ReDim m_Pixels(0 To imgWidth - 1, 0 To imgHeight - 1) As Byte
                                End If
                                
                                If m_Resources.GetNodeDataByIndex(nodeIndex, False, m_Buffer, False) Then
                                    
                                    'Copy the palette and pixel data into place
                                    CopyMemory ByVal VarPtr(m_Palette(0)), ByVal VarPtr(m_Buffer(0)), numColors * 4
                                    CopyMemory ByVal VarPtr(m_Pixels(0, 0)), ByVal VarPtr(m_Buffer(numColors * 4)), imgWidth * imgHeight
                                    
                                    'Build a matching 32-bpp DIB from the palette and pixel data
                                    LoadImageResource = DIBs.GetRGBADIB_FromPalette(m_tmpDIB, numColors, m_Palette, m_Pixels)
                                
                                End If
                                
                            End If
                        
                        Else
                            Dim tmpDIBPointer As Long, tmpDIBLength As Long
                            m_tmpDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                            LoadImageResource = m_Resources.GetNodeDataByIndex_UnsafeDstPointer(nodeIndex, False, tmpDIBPointer)
                        End If
                        
                    End If
                    
                    'At present, all resources contain premultiplied alpha, so force the corresponding state now
                    m_tmpDIB.SetInitialAlphaPremultiplicationState True
                    
                    'Resize the DIB into the destination
                    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
                    If (desiredWidth = 0) Then desiredWidth = 16#
                    If (desiredHeight = 0) Then desiredHeight = 16#
                    
                    Dim desiredWidthInt As Long, desiredHeightInt As Long
                    desiredWidthInt = Int(desiredWidth + 0.9999)
                    desiredHeightInt = Int(desiredHeight + 0.9999)
                    
                    If (dstDIB.GetDIBWidth <> desiredWidthInt) Or (dstDIB.GetDIBHeight <> desiredHeightInt) Then
                        dstDIB.CreateBlank desiredWidthInt, desiredHeightInt, 32, 0, 0
                    Else
                        dstDIB.ResetDIB 0
                    End If
                    
                    dstDIB.SetInitialAlphaPremultiplicationState True
                    GDI_Plus.GDIPlus_StretchBlt dstDIB, desiredBorders, desiredBorders, desiredWidth - desiredBorders * 2, desiredHeight - desiredBorders * 2, m_tmpDIB, 0, 0, imgWidth, imgHeight, , GP_IM_HighQualityBicubic, , True, , True
                    
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
                            If m_XML.GetUniqueTag_Boolean("rt-clr") Then
                            
                                'Retrieve the image's color.  (NOTE: the specified color will be overridden with
                                ' monochrome if the monochrome icon preference is active; see the branch above.)
                                If dstIsMenu Then
                                    If m_XML.GetUniqueTag_Boolean("rt-clrmenu", False) Then
                                        targetColor = m_XML.GetUniqueTag_Long("clr-m")
                                    Else
                                        targetColor = m_XML.GetUniqueTag_Long("clr-l")
                                    End If
                                Else
                                    If ((g_Themer.GetCurrentThemeClass = PDTC_Light) Or dstIsMenu) Then
                                        targetColor = m_XML.GetUniqueTag_Long("clr-l")
                                    ElseIf (g_Themer.GetCurrentThemeClass = PDTC_Dark) Then
                                        targetColor = m_XML.GetUniqueTag_Long("clr-d")
                                    End If
                                End If
                                
                                DIBs.ColorizeDIB dstDIB, targetColor
                            
                            End If
                            
                        End If
                        
                    End If
                    
                    'Before returning, free the target DIB from its DC.  (There's a chance our caller won't
                    ' use the DIB right away, and this keeps an unnecessary GDI object from being created.)
                    dstDIB.FreeFromDC
                    
                End If
                
            End If
            
        End If
    
    End If
    
End Function

Private Sub RaiseResourceError(ByVal msgError As String, Optional ByVal msgIsNonErrorFeedback As Boolean = False)
    #If DEBUGMODE = 1 Then
        If msgIsNonErrorFeedback Then
            pdDebug.LogAction "pdResources reported: " & msgError
        Else
            pdDebug.LogAction "WARNING!  pdResources error: " & msgError
        End If
    #End If
End Sub

Private Sub Class_Initialize()
    
    m_ResourcesAvailable = False
    m_ThemeIconsMonochrome = False
    
    Set m_XML = New pdXML
    m_XML.SetTextCompareMode vbBinaryCompare
    
End Sub
