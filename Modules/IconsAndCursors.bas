Attribute VB_Name = "IconsAndCursors"
'***************************************************************************
'PhotoDemon Icon and Cursor Handler
'Copyright 2012-2018 by Tanner Helland
'Created: 24/June/12
'Last updated: 03/February/17
'Last update: add automatic high-DPI cursor support for custom cursors created from PNGs
'
'Because VB6 doesn't provide many mechanisms for working with icons, I've had to manually add a number of icon-related
' functions to PhotoDemon.  As of 7.0, all icons in the program are stored in PD's custom resource file (in a variety of
' formats, each one optimized for file size).  These icons are extracted, re-colored, resized (for high-DPI screens),
' and rendered onto UI elements at run-time.
'
'Menu icons currently use the clsMenuImage class by Leandro Ascierto.  Please see that class for details on how it works.
' (A link to Leandro's original project can also be found there.)
'
'This module also handles the rendering of dynamic form, program, and taskbar icons.  When an image is loaded and active,
' those icons can change to match the current image.  For an overview on how this works, you can see the following MSDN page:
' http://support.microsoft.com/kb/318876
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'API calls for building an icon at run-time
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal cPlanes As Long, ByVal cBitsPerPel As Long, ByVal lpvBits As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (icoInfo As ICONINFO) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long

'System constants for retrieving system default icon sizes and related metrics
Private Const SM_CXICON As Long = 11
Private Const SM_CYICON As Long = 12
Private Const SM_CXSMICON As Long = 49
Private Const SM_CYSMICON As Long = 50
Private Const LR_SHARED As Long = &H8000&
Private Const IMAGE_ICON As Long = 1
Private Const WM_SETICON As Long = &H80
Private Const ICON_SMALL As Long = 0
Private Const ICON_BIG As Long = 1

'API needed for converting PNG data to icon or cursor format
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef mImage As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef hbmReturn As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef hBmpReturn As Long, ByVal Background As Long) As GP_Result
Private Declare Function GdipGetImageBounds Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef mSrcRect As RectF, ByRef mSrcUnit As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal gdiBitmap As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef PixelFormat As Long) As Long

'Used to check GDI+ images for alpha channels
Private Const PixelFormatAlpha = &H40000             ' Has an alpha component
Private Const PixelFormatPAlpha = &H80000            ' Pre-multiplied alpha

'GDI+ types and constants
Private Const UnitPixel As Long = &H2&

'Type required to create an icon on-the-fly
Private Type ICONINFO
    fIcon As Boolean
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

'Used to apply and manage custom cursors (without subclassing)
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long

Public Enum SystemCursorConstant
    IDC_DEFAULT = 0&
    IDC_APPSTARTING = 32650&
    IDC_HAND = 32649&
    IDC_ARROW = 32512&
    IDC_CROSS = 32515&
    IDC_IBEAM = 32513&
    IDC_ICON = 32641&
    IDC_NO = 32648&
    IDC_SIZEALL = 32646&
    IDC_SIZENESW = 32643&
    IDC_SIZENS = 32645&
    IDC_SIZENWSE = 32642&
    IDC_SIZEWE = 32644&
    IDC_UPARROW = 32516&
    IDC_WAIT = 32514&
End Enum

#If False Then
    Private Const IDC_DEFAULT = 0&, IDC_APPSTARTING = 32650&, IDC_HAND = 32649&, IDC_ARROW = 32512&, IDC_CROSS = 32515&, IDC_IBEAM = 32513&, IDC_ICON = 32641&, IDC_NO = 32648&, IDC_SIZEALL = 32646&, IDC_SIZENESW = 32643&, IDC_SIZENS = 32645&, IDC_SIZENWSE = 32642&, IDC_SIZEWE = 32644&, IDC_UPARROW = 32516&, IDC_WAIT = 32514&
#End If

Private Const GCL_HCURSOR = (-12)

Private m_numCustomCursors As Long
Private m_customCursorNames() As String
Private m_customCursorHandles() As Long

'This array will be used to store our dynamically created icon handles so we can delete them on program exit
Private Const INITIAL_ICON_CACHE_SIZE As Long = 16
Private m_numOfIcons As Long
Private m_iconHandles() As Long

'As of v7.0, icon creation and destruction is tracked locally.
Private m_IconsCreated As Long, m_IconsDestroyed As Long

'This constant is used for testing only.  It should always be set to TRUE for production code.
Private Const ALLOW_DYNAMIC_ICONS As Boolean = True

'This array tracks the resource identifiers and consequent numeric identifiers of all loaded icons.  The size of the array
' is arbitrary; just make sure it's larger than the max number of loaded icons.
Private iconNames(0 To 511) As String

'We also need to track how many icons have been loaded; this counter will also be used to reference icons in the database
Private curIcon As Long

'clsMenuImage does the heavy lifting for inserting icons into menus
Private cMenuImage As clsMenuImage

'A second class is used to manage the icons for the MRU list.
Private cMRUIcons As clsMenuImage

'PD's default large and small application icons.  These are cached for the duration of the current session.
Private m_DefaultIconLarge As Long, m_DefaultIconSmall As Long

'Size of the icons in the "Recent Files" menu.  This size is larger on Vista+ because it allows each menu to have its
' own image size (vs XP, where they must all match).
Private m_RecentFileIconSize As Long

'System cursor size (x, y).  X and Y sizes should always be identical, so we only cache one.
Private m_CursorSize As Long

'Load all the menu icons from PhotoDemon's embedded resource file
Public Sub LoadMenuIcons(Optional ByVal alsoApplyMenuIcons As Boolean = True)

    FreeMenuIconCache
    
    With cMenuImage
            
        'Use Leandro's class to check if the current Windows install supports theming.
        g_IsThemingEnabled = .CanWeTheme
    
        'Disable menu icon drawing if on Windows XP and uncompiled (to prevent subclassing crashes on unclean IDE breaks)
        If (Not OS.IsVistaOrLater) And (Not OS.IsProgramCompiled) Then
            Debug.Print "XP + IDE detected.  Menu icons will be disabled for this session."
            Exit Sub
        End If
        
        .Init FormMain.hWnd, FixDPI(16), FixDPI(16)
        
    End With
            
    'Now that all menu icons are loaded, apply them to the proper menu entires
    If alsoApplyMenuIcons Then IconsAndCursors.ApplyAllMenuIcons
        
    '...and initialize the separate MRU icon handler.
    Set cMRUIcons = New clsMenuImage
    If OS.IsVistaOrLater Then m_RecentFileIconSize = FixDPI(64) Else m_RecentFileIconSize = FixDPI(16)
    cMRUIcons.Init FormMain.hWnd, m_RecentFileIconSize, m_RecentFileIconSize
        
End Sub

Public Sub FreeMenuIconCache()
    
    'If we are re-loading all icons instead of just loading them for the first time, clear out the old list
    If (Not cMenuImage Is Nothing) Then
        cMenuImage.Clear
        Set cMenuImage = Nothing
    End If
    
    Set cMenuImage = New clsMenuImage
    
    'Also reset the icon tracking array
    curIcon = 0
    Erase iconNames
    
End Sub

'Apply (and if necessary, dynamically load) menu icons to their proper menu entries.
Public Sub ApplyAllMenuIcons()
    Menus.ApplyIconsToMenus
End Sub

'This new, simpler technique for adding menu icons requires only the menu location (including sub-menus) and the icon's identifer
' in the resource file.  If the icon has already been loaded, it won't be loaded again; instead, the function will check the list
' of loaded icons and automatically fill in the numeric identifier as necessary.
Public Sub AddMenuIcon(ByRef resID As String, ByVal topMenu As Long, ByVal subMenu As Long, Optional ByVal subSubMenu As Long = -1)
    
    If (cMenuImage Is Nothing) Then Exit Sub
    
    On Error GoTo MenuIconNotFound
    
    Dim i As Long
    Dim iconLocation As Long
    Dim iconAlreadyLoaded As Boolean
    
    iconAlreadyLoaded = False
    
    'Loop through all icons that have been loaded, and see if this one has been requested already.
    ' (This is necessary because some menus reuse the same icons, and it's a waste of resources to maintain
    '  two copies of said icons.)
    For i = 0 To curIcon

        If Strings.StringsEqual(iconNames(i), resID, False) Then
            iconAlreadyLoaded = True
            iconLocation = i
            Exit For
        End If

    Next i
    
    'If the icon was not found, load it and add it to the list
    If (Not iconAlreadyLoaded) Then
        AddImageResourceToClsMenu resID, cMenuImage
        iconNames(curIcon) = resID
        iconLocation = curIcon
        curIcon = curIcon + 1
    End If
        
    'Place the icon onto the requested menu
    If (subSubMenu = -1) Then
        cMenuImage.PutImageToVBMenu iconLocation, subMenu, topMenu
    Else
        cMenuImage.PutImageToVBMenu iconLocation, subSubMenu, topMenu, subMenu
    End If
    
MenuIconNotFound:

End Sub

'When menu captions are changed, the associated images are lost.  This forces a reset.
' Note that to keep the code small, all changeable icons are refreshed whenever this is called.
Public Sub ResetMenuIcons()
    
    'Redraw the Undo/Redo menus
    AddMenuIcon "edit_undo", 1, 0     'Undo
    AddMenuIcon "edit_redo", 1, 1     'Redo
    
    'Redraw the Repeat and Fade menus
    AddMenuIcon "edit_repeat", 1, 4         'Repeat previous action
    
    'Redraw the Window menu, as some of its menus will be en/disabled according to the docking status of image windows
    AddMenuIcon "generic_next", 9, 7       'Next image
    AddMenuIcon "generic_previous", 9, 8   'Previous image
    
    'Dynamically calculate the position of the Clear Recent Files menu item and update its icon
    Dim numOfMRUFiles As Long
    If (Not g_RecentFiles Is Nothing) Then
        
        'Start by making sure all menu captions are correct
        Menus.UpdateSpecialMenu_RecentFiles
        
        'Retrieve the number of files displayed in the menu.  (Note that this is *not* the same number
        ' as the menu count, as we add some extra options after the MRU entries themselves.)
        numOfMRUFiles = g_RecentFiles.GetNumOfItems()
        
    End If
    
    'Clear the current MRU icon cache.
    If (Not cMRUIcons Is Nothing) Then
        
        cMRUIcons.Clear
        
        If (numOfMRUFiles > 0) Then
        
            'Load a placeholder image for missing MRU entries
            AddImageResourceToClsMenu "generic_imagemissing", cMRUIcons, m_RecentFileIconSize
            
            'This counter will be used to track the current position of loaded thumbnail images into the icon collection
            Dim iconLocation As Long
            iconLocation = 0
            
            Dim tmpDIB As pdDIB: Set tmpDIB = New pdDIB
            
            'Loop through the MRU list, and attempt to load thumbnail images for each entry
            Dim i As Long
            For i = 0 To numOfMRUFiles - 1
            
                'Start by seeing if an image exists for this MRU entry
                If (g_RecentFiles.GetMRUThumbnail(i) Is Nothing) Then
                    
                    'If a thumbnail doesn't exist, supply a placeholder image (Vista+ only; on XP it will simply be blank)
                    If OS.IsVistaOrLater Then cMRUIcons.PutImageToVBMenu 0, i, 0, 2
                
                'A thumbnail exists.  Load it directly from the source DIB.
                Else
                
                    'Note that a temporary DIB is required, because we transfer ownership of the DIB to the menu manager.
                    ' (If we copy the existing DIB as-is, it will be removed from the MRU collection, which we definitely
                    ' don't want as we need to write that DIB out to file when the program closes!)
                    tmpDIB.CreateFromExistingDIB g_RecentFiles.GetMRUThumbnail(i)
                    iconLocation = iconLocation + 1
                    cMRUIcons.AddImageFromDIB tmpDIB
                    cMRUIcons.PutImageToVBMenu iconLocation, i, 0, 2
                    
                End If
                
            Next i
                
            'Vista+ users now get their nice, large "load all recent files" and "clear list" icons.
            If OS.IsVistaOrLater Then
            
                Dim largePadding As Single
                largePadding = (m_RecentFileIconSize * 0.2)
                AddImageResourceToClsMenu "generic_imagefolder", cMRUIcons, m_RecentFileIconSize, largePadding
                cMRUIcons.PutImageToVBMenu iconLocation + 1, numOfMRUFiles + 1, 0, 2
                largePadding = (m_RecentFileIconSize * 0.333)
                AddImageResourceToClsMenu "file_close", cMRUIcons, m_RecentFileIconSize, largePadding
                cMRUIcons.PutImageToVBMenu iconLocation + 2, numOfMRUFiles + 2, 0, 2
            
            'XP users are stuck with little 16x16 ones (to match the rest of the menu icons)
            Else
                AddMenuIcon "generic_imagefolder", 0, 2, numOfMRUFiles + 1
                AddMenuIcon "file_close", 0, 2, numOfMRUFiles + 2
            End If
        
        'When the current list is empty, we display an icon-less "Empty" statement
        Else
            cMRUIcons.PutImageToVBMenu -1, 0, 0, 2
        End If
        
    End If
    
    'Repeat the same steps for the Recent Macro list.  Note that a larger icon is never used for this list,
    ' because we don't display thumbnail images for macros (so just a default 16x16 icon is acceptable).
    If (Not g_RecentMacros Is Nothing) Then
        
        'Again, make sure all menu captions are correct
        Menus.UpdateSpecialMenu_RecentMacros
        
        Dim numOfMRUFiles_Macro As Long
        numOfMRUFiles_Macro = g_RecentMacros.MRU_ReturnCount
        If (numOfMRUFiles_Macro > 0) Then
            AddMenuIcon "file_close", 8, 7, numOfMRUFiles_Macro + 1
        
        'If the list is empty, do not display any icons
        Else
            AddMenuIcon "file_close", 8, 7, 2
            If (Not cMRUIcons Is Nothing) Then cMRUIcons.PutImageToVBMenu -1, 0, 8, 7
        End If
        
    End If
        
End Sub

Private Sub AddImageResourceToClsMenu(ByRef srcResID As String, ByRef targetMenuObject As clsMenuImage, Optional ByVal desiredSize As Long = 0, Optional ByVal desiredPadding As Single = 0.5)

    'First, attempt to load the image from our internal resource manager
    Dim loadedInternally As Boolean: loadedInternally = False
    If (Not g_Resources Is Nothing) Then
        If g_Resources.AreResourcesAvailable Then
            Dim tmpDIB As pdDIB
            If (desiredSize = 0) Then desiredSize = FixDPI(16)
            loadedInternally = g_Resources.LoadImageResource(srcResID, tmpDIB, desiredSize, desiredSize, desiredPadding, True)
            If loadedInternally Then targetMenuObject.AddImageFromDIB tmpDIB
        End If
    End If
    
    'If that fails, use the legacy resource instead
    If (Not loadedInternally) Then
        targetMenuObject.AddImageFromStream LoadResData(srcResID, "CUSTOM")
    End If
    
End Sub

'Convert a DIB - any DIB! - to an icon via CreateIconIndirect.  Transparency will be preserved, and by default, the icon will be created
' at the current image's size (though you can specify a custom size if you wish).  Ideally, the passed DIB will have been created using
' the pdImage function "RequestThumbnail".
'
'FreeImage is currently required for this function, because it provides a simple way to move between DIBs and DDBs.  I could rewrite
' the function without FreeImage's help, but frankly don't consider it worth the trouble.
Public Function GetIconFromDIB(ByRef srcDIB As pdDIB, Optional iconSize As Long = 0) As Long

    'If the iconSize parameter is 0, use the current DIB's dimensions.  Otherwise, resize it as requested.
    If (iconSize = 0) Then iconSize = srcDIB.GetDIBWidth
    
    'If the requested icon size does not match the incoming DIB's size, we need to create a temporary DIB at the correct size.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    If (iconSize <> srcDIB.GetDIBWidth) Or (iconSize <> srcDIB.GetDIBHeight) Then
        
        tmpDIB.CreateBlank iconSize, iconSize, 32, 0, 0
        
        'To improve quality at very low sizes, enforce prefiltering
        Dim resampleMode As GP_InterpolationMode
        If (iconSize <= 32) Then resampleMode = GP_IM_HighQualityBicubic Else resampleMode = GP_IM_HighQualityBicubic
        GDI_Plus.GDIPlus_StretchBlt tmpDIB, 0, 0, iconSize, iconSize, srcDIB, 0, 0, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, , g_UserPreferences.GetThumbnailInterpolationPref(), , , , True
        
    Else
        tmpDIB.CreateFromExistingDIB srcDIB
    End If
    
    'tmpDIB now points at a properly resized source DIB object.  If the DIB is 32-bpp, unpremultply the alpha now.
    If (tmpDIB.GetDIBColorDepth = 32) And tmpDIB.GetAlphaPremultiplication Then tmpDIB.SetAlphaPremultiplication False
    
    'Icon generation has a number of quirks.  One is that even if you want a 32bpp icon, you still must supply a blank
    ' monochrome mask for the icon, even though the API just discards it.  Prepare such a mask now.
    Dim monoBmp As Long
    monoBmp = CreateBitmap(iconSize, iconSize, 1&, 1&, ByVal 0&)
    
    'Create a header for the icon we desire, then use CreateIconIndirect to create it.
    Dim icoInfo As ICONINFO
    With icoInfo
        .fIcon = True
        .xHotspot = iconSize
        .yHotspot = iconSize
        .hbmMask = monoBmp
        .hbmColor = GDI.GetDDBFromDIB(tmpDIB)
    End With
        
    GetIconFromDIB = CreateNewIcon(icoInfo)
    
    'Delete the temporary monochrome mask and DDB
    DeleteObject monoBmp
    DeleteObject icoInfo.hbmColor
    
End Function

'Create a custom form icon for a target pdImage object
Public Sub CreateCustomFormIcons(ByRef srcImage As pdImage)

    If (ALLOW_DYNAMIC_ICONS And (Not srcImage Is Nothing)) Then
    
        Dim thumbDIB As pdDIB
        
        'Request a 32x32 thumbnail version of the current image
        If srcImage.RequestThumbnail(thumbDIB, 32) Then
            
            'Request two icon-format versions of the generated thumbnail.
            ' (Taskbar icons are generally 32x32.  Form titlebar icons are generally 16x16.)
            Dim hIcon32 As Long, hIcon16 As Long
            hIcon32 = GetIconFromDIB(thumbDIB)
            hIcon16 = GetIconFromDIB(thumbDIB, 16)
            
            'Each pdImage instance caches its custom icon handles, which simplifies the process of synchronizing PD's icons
            ' to any given image if the user is working with multiple images at once.  Retrieve the old handles now, so we
            ' can free them after we set the new ones.
            Dim oldIcon32 As Long, oldIcon16 As Long
            oldIcon32 = srcImage.GetImageIcon(True)
            oldIcon16 = srcImage.GetImageIcon(False)
            
            'Set the new icons, then free the old ones
            srcImage.SetImageIcon True, hIcon32
            srcImage.SetImageIcon False, hIcon16
            If (oldIcon32 <> 0) Then ReleaseIcon oldIcon32
            If (oldIcon16 <> 0) Then ReleaseIcon oldIcon16
            
        Else
            Debug.Print "WARNING!  Image refused to provide a thumbnail!"
        End If
        
    End If

End Sub

Private Function CreateNewIcon(ByRef icoStruct As ICONINFO) As Long
    CreateNewIcon = CreateIconIndirect(icoStruct)
    If ((CreateNewIcon <> 0) And icoStruct.fIcon) Then m_IconsCreated = m_IconsCreated + 1
End Function

Public Sub ReleaseIcon(ByVal hIcon As Long)
    If (hIcon <> 0) Then
        DestroyIcon hIcon
        m_IconsDestroyed = m_IconsDestroyed + 1
    End If
End Sub

Public Function GetCreatedIconCount(Optional ByRef iconsCreated As Long, Optional ByRef iconsDestroyed As Long) As Long
    iconsCreated = m_IconsCreated
    iconsDestroyed = m_IconsDestroyed
    GetCreatedIconCount = m_IconsCreated - m_IconsDestroyed
End Function

'Needs to be run only once, at the start of the program
Public Sub InitializeIconHandler()
    m_numOfIcons = 0
    ReDim m_iconHandles(0 To INITIAL_ICON_CACHE_SIZE - 1) As Long
End Sub

Private Sub AddIconToList(ByVal hIcon As Long)
    
    If (m_numOfIcons > UBound(m_iconHandles)) Then ReDim Preserve m_iconHandles(0 To UBound(m_iconHandles) * 2 + 1) As Long
    
    m_iconHandles(m_numOfIcons) = hIcon
    m_numOfIcons = m_numOfIcons + 1

End Sub

'Remove all icons generated since the program launched
Public Sub DestroyAllIcons()

    If (m_numOfIcons > 0) Then
    
        Dim i As Long
        For i = 0 To m_numOfIcons - 1
            If (m_iconHandles(i) <> 0) Then ReleaseIcon m_iconHandles(i)
        Next i
        
        'Reinitialize the icon handler, which will also reset the icon count and handle array
        InitializeIconHandler
        
    End If

End Sub

'Given an image in the .exe's resource section (typically a PNG image), return an icon handle to it (hIcon).
' The calling function is responsible for deleting this object once they are done with it.
Public Function CreateIconFromResource(ByVal resTitle As String) As Long
    
    'Start by extracting the PNG data into a bytestream
    Dim imageData() As Byte
    imageData() = LoadResData(resTitle, "CUSTOM")
    
    Dim hBitmap As Long, hIcon As Long
    
    Dim IStream As IUnknown
    Set IStream = VBHacks.GetStreamFromVBArray(VarPtr(imageData(0)), UBound(imageData) - LBound(imageData) + 1)
    
    If Not (IStream Is Nothing) Then
        
        'Note that GDI+ will have been initialized already, as part of the core PhotoDemon startup routine
        If (GdipLoadImageFromStream(IStream, hBitmap) = 0) Then
        
            'hBitmap now contains the PNG file as an hBitmap (obviously).  Now we need to convert it to icon format.
            If (GdipCreateHICONFromBitmap(hBitmap, hIcon) = 0) Then
                CreateIconFromResource = hIcon
            Else
                CreateIconFromResource = 0
            End If
            
            GdipDisposeImage hBitmap
                
        End If
    
        Set IStream = Nothing
    
    End If
    
    Exit Function
    
End Function

'Given an image in PD's theme file, return it as a system cursor handle.
'PLEASE NOTE: PD auto-caches created cursors, and retains them for the life of the program.  They should not be manually freed.
Public Function CreateCursorFromResource(ByVal resTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0) As Long
    
    'Ensure our cached system cursor size is up-to-date
    GetSystemCursorSizeInPx
    
    'Start by extracting the image itself into a DIB.  Note that the image will be auto-scaled to the system
    ' cursor size calculated above.
    Dim resDIB As pdDIB
    Set resDIB = New pdDIB
    If LoadResourceToDIB(resTitle, resDIB, m_CursorSize, m_CursorSize, , , True) Then
        
        'Next, we need to scale the cursor hotspot to the actual cursor size.
        ' (Cursor hotspots are always hard-coded on a 16x16 basis, then adjusted at run-time as necessary.)
        curHotspotX = (CDbl(curHotspotX) / 16#) * m_CursorSize
        curHotspotY = (CDbl(curHotspotY) / 16#) * m_CursorSize
        
        'Generate a blank monochrome mask to pass to the icon creation function.
        ' (This is a stupid Windows workaround for 32bpp cursors.  The cursor creation function always assumes
        '  the presence of a mask bitmap, so we have to submit one even if we want the PNG's alpha channel
        '  used for transparency.)
        Dim monoBmp As Long
        monoBmp = CreateBitmap(resDIB.GetDIBWidth, resDIB.GetDIBHeight, 1, 1, ByVal 0&)
        
        'Create an icon header and point it at our temporary mask and original DIB resource
        Dim icoInfo As ICONINFO
        With icoInfo
            .fIcon = False
            .xHotspot = curHotspotX
            .yHotspot = curHotspotY
            .hbmMask = monoBmp
            .hbmColor = resDIB.GetDIBHandle
        End With
        
        'Create the cursor
        CreateCursorFromResource = CreateNewIcon(icoInfo)
        
        'Release our temporary mask and resource container, as Windows has now made its own copies
        DeleteObject monoBmp
        Set resDIB = Nothing
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  IconsAndCursors.CreateCursorFromResource failed to find the resource: " & resTitle
        #End If
    End If
    
    Exit Function
    
End Function

'Retrieve the current system cursor size, in pixels.  Make sure to read the function details - the size returned differs
' from what bare WAPI functions return, by design.
Public Function GetSystemCursorSizeInPx() As Long

    'If this is the first custom cursor request, cache the current system cursor size.
    ' (This function matches those sizes automatically, and the caller does not have control over it, by design.)
    
    'Also, note that Windows cursors typically only use one quadrant of the current system cursor size.  This odd behavior
    ' is why we divide the retrieved cursor size by two.
    If (m_CursorSize = 0) Then
        Const SM_CYCURSOR As Long = 14
        m_CursorSize = GetSystemMetrics(SM_CYCURSOR) \ 2
        If (m_CursorSize <= 0) Then m_CursorSize = FixDPI(16)
    End If
    
    GetSystemCursorSizeInPx = m_CursorSize

End Function

'Load all relevant program cursors into memory
Public Sub InitializeCursors()

    ReDim m_customCursorHandles(0) As Long

    'Previously, system cursors were cached here.  This is no longer needed per https://github.com/tannerhelland/PhotoDemon/issues/78
    ' I am leaving this sub in case I need to pre-load tool cursors in the future.
    
    'Note that UnloadAllCursors below is still required, as the program may dynamically generate custom cursors while running, and
    ' these cursors will not be automatically deleted by the system.

End Sub

'Unload any custom cursors from memory
Public Sub UnloadAllCursors()
    
    If (m_numCustomCursors = 0) Then Exit Sub
    
    Dim i As Long
    For i = 0 To m_numCustomCursors - 1
        DestroyCursor m_customCursorHandles(i)
    Next i
    
End Sub

'Use any 32bpp PNG resource as a cursor .  When setting the mouse pointer of VB objects, please use SetPNGCursorToObject, below.
Public Sub SetPNGCursorToHwnd(ByVal dstHwnd As Long, ByVal pngTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0)
    SetClassLong dstHwnd, GCL_HCURSOR, RequestCustomCursor(pngTitle, curHotspotX, curHotspotY)
End Sub

'Use any 32bpp PNG resource as a cursor.  Use this function preferentially over the previous one, "SetPNGCursorToHwnd", when possible.
' (If a VB object does not have its MousePointer property set to "custom", it will override our attempts to set a custom mouse icon.)
Public Sub SetPNGCursorToObject(ByRef srcObject As Object, ByVal pngTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0)
    srcObject.MousePointer = vbCustom
    SetClassLong srcObject.hWnd, GCL_HCURSOR, RequestCustomCursor(pngTitle, curHotspotX, curHotspotY)
End Sub

'Set a single object to use the hand cursor
Public Sub SetHandCursor(ByRef tControl As Object)
    tControl.MouseIcon = LoadPicture(vbNullString)
    tControl.MousePointer = 99
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_HAND)
End Sub

Public Sub SetHandCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, LoadCursor(0, IDC_HAND)
End Sub

Public Sub SetArrowCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, LoadCursor(0, IDC_ARROW)
End Sub

'Set a single form to use the arrow cursor
Public Sub SetArrowCursor(ByRef tControl As Object)
    tControl.MousePointer = vbCustom
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_ARROW)
End Sub

'If a custom PNG cursor has not been loaded, this function will load the PNG, convert it to cursor format, then store
' the cursor resource for future reference (so the image doesn't have to be loaded again).
Public Function RequestCustomCursor(ByVal resCursorName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0) As Long

    Dim i As Long
    Dim cursorLocation As Long
    Dim cursorAlreadyLoaded As Boolean
    
    cursorLocation = 0
    cursorAlreadyLoaded = False
    
    'Loop through all cursors that have been loaded, and see if this one has been requested already.
    If (m_numCustomCursors > 0) Then
    
        For i = 0 To m_numCustomCursors - 1
        
            If Strings.StringsEqual(m_customCursorNames(i), resCursorName, False) Then
                cursorAlreadyLoaded = True
                cursorLocation = i
                Exit For
            End If
        
        Next i
    
    End If
    
    'If the cursor was not found, load it and add it to the list
    If cursorAlreadyLoaded Then
        RequestCustomCursor = m_customCursorHandles(cursorLocation)
    Else
        Dim tmpHandle As Long
        tmpHandle = CreateCursorFromResource(resCursorName, cursorHotspotX, cursorHotspotY)
        
        If (tmpHandle <> 0) Then
            ReDim Preserve m_customCursorNames(0 To m_numCustomCursors) As String
            ReDim Preserve m_customCursorHandles(0 To m_numCustomCursors) As Long
            m_customCursorNames(m_numCustomCursors) = resCursorName
            m_customCursorHandles(m_numCustomCursors) = tmpHandle
            m_numCustomCursors = m_numCustomCursors + 1
        End If
        
        RequestCustomCursor = tmpHandle
    End If

End Function

'Given an image in the .exe's resource section (typically a PNG image), load it to a pdDIB object.
' The calling function is responsible for deleting the DIB once they are done with it.
Public Function LoadResourceToDIB(ByVal resTitle As String, ByRef dstDIB As pdDIB, Optional ByVal desiredWidth As Long = 0, Optional ByVal desiredHeight As Long = 0, Optional ByVal desiredBorders As Long = 0, Optional ByVal useCustomColor As Long = -1, Optional ByVal suspendMonochrome As Boolean = False, Optional ByVal resampleAlgorithm As GP_InterpolationMode = GP_IM_HighQualityBicubic) As Boolean
        
    'As of v7.0, PD now has two places from which to pull resources:
    ' 1) Its own custom resource handler (which is the preferred location)
    ' 2) The old, standard .exe resource section (which is deprecated, and in the process of being removed)
    '
    'We always attempt (1) before falling back to (2).  The goal for 7.0's release is to remove (2) entirely.
        
    'Some functions may call this before GDI+ has loaded; exit if that happens
    If Drawing2D.IsRenderingEngineActive(P2_GDIPlusBackend) Then
    
        'Attempt the default resource manager first
        Dim intResFound As Boolean: intResFound = False
        If (Not g_Resources Is Nothing) Then
            If g_Resources.AreResourcesAvailable Then
            
                'Attempt to load the requested resource.  (This may fail, as I am still in the process of migrating
                ' all resources to the new format.)
                intResFound = g_Resources.LoadImageResource(resTitle, dstDIB, desiredWidth, desiredHeight, desiredBorders, , useCustomColor, suspendMonochrome, resampleAlgorithm)
                LoadResourceToDIB = intResFound
            
            End If
        End If
        
        'If we failed to find the requested resource, return a small blank DIB (to work around some old code
        ' that expects an initialized DIB), report the problem, then exit
        If (Not intResFound) Then
        
            If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
            dstDIB.CreateBlank 16, 16, 32, 0, 0
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  LoadResourceToDIB couldn't find <" & resTitle & ">.  Check your spelling and try again."
            #End If
            
            LoadResourceToDIB = False
            
        End If
        
    Else
        'Debug.Print "GDI+ unavailable; resources suspended for this session."
        LoadResourceToDIB = False
    End If
    
End Function

'PD will automatically update its taskbar icon to reflect the current image being edited.  I find this especially helpful
' when multiple PD sessions are operating in parallel.
Public Sub ChangeAppIcons(ByVal hIconSmall As Long, ByVal hIconLarge As Long)
    
    If (Not ALLOW_DYNAMIC_ICONS) Then Exit Sub
    Dim oldHIconL As Long, oldHIconS As Long
    oldHIconS = SendMessageA(FormMain.hWnd, WM_SETICON, ICON_SMALL, ByVal hIconSmall)
    oldHIconL = SendMessageA(FormMain.hWnd, WM_SETICON, ICON_BIG, ByVal hIconLarge)
    
    'Generally speaking, you want to destroy the old icons after a change, but we track (and manage)
    ' these values internally, so there's no need to destroy icons at WM_SETICON time.
    'If (oldHIconS <> 0) Then DestroyIcon oldHIconS
    'If (oldHIconL <> 0) Then DestroyIcon oldHIconL
    
End Sub

'When loading a modal dialog, the dialog will not have an icon by default.  We can assign an icon at run-time to ensure that icons
' appear in the Alt+Tab dialog of older OSes.
Public Sub ChangeWindowIcon(ByVal targetHWnd As Long, ByVal hIconSmall As Long, ByVal hIconLarge As Long, Optional ByRef dstSmallIcon As Long = 0, Optional ByRef dstLargeIcon As Long = 0)
    If (targetHWnd <> 0) Then
        dstSmallIcon = SendMessageA(targetHWnd, WM_SETICON, ICON_SMALL, ByVal hIconSmall)
        dstLargeIcon = SendMessageA(targetHWnd, WM_SETICON, ICON_BIG, ByVal hIconLarge)
    End If
End Sub

Public Sub MirrorCurrentIconsToWindow(ByVal targetHWnd As Long, Optional ByVal setLargeIconOnly As Boolean = False, Optional ByRef dstSmallIcon As Long = 0, Optional ByRef dstLargeIcon As Long = 0)
    If (g_OpenImageCount > 0) Then
        ChangeWindowIcon targetHWnd, IIf(setLargeIconOnly, 0&, pdImages(g_CurrentImage).GetImageIcon(False)), pdImages(g_CurrentImage).GetImageIcon(True), dstSmallIcon, dstLargeIcon
    Else
        ChangeWindowIcon targetHWnd, IIf(setLargeIconOnly, 0&, m_DefaultIconSmall), m_DefaultIconLarge, dstSmallIcon, dstLargeIcon
    End If
End Sub

'When all images are unloaded (or when the program is first loaded), we must reset the program icon to its default values.
Public Sub ResetAppIcons()
    
    If (m_DefaultIconLarge = 0) Then
        m_DefaultIconLarge = LoadImageAsString(App.hInstance, "AAA", IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_SHARED)
    End If
    
    If (m_DefaultIconSmall = 0) Then
        m_DefaultIconSmall = LoadImageAsString(App.hInstance, "AAA", IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_SHARED)
    End If
    
    ChangeAppIcons m_DefaultIconSmall, m_DefaultIconLarge
    
End Sub

'When PD is first loaded, we associate an icon with the master "ThunderMain" owner window, to ensure proper icons in places
' like Task Manager.
Public Sub SetThunderMainIcon()

    'Start by loading the default icons from the resource file, as necessary
    ResetAppIcons
    
    Dim tmHWnd As Long
    tmHWnd = OS.ThunderMainHWnd()
    SendMessageA tmHWnd, WM_SETICON, ICON_SMALL, ByVal m_DefaultIconLarge
    SendMessageA tmHWnd, WM_SETICON, ICON_BIG, ByVal m_DefaultIconSmall

End Sub
