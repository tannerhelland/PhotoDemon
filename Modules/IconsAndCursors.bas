Attribute VB_Name = "IconsAndCursors"
'***************************************************************************
'PhotoDemon Icon and Cursor Handler
'Copyright 2012-2026 by Tanner Helland
'Created: 24/June/12
'Last updated: 14/December/21
'Last update: fix cursor hotspot placement bug when creating cursors from resources
'
'Because VB6 doesn't provide many (any?) mechanisms for manipulating icons, I've had to manually write a wide variety
' of icon handling functions.  As of v7.0, all icons in PD are stored in our custom resource file (in a variety of
' formats, each one optimized for file size).  These icons are extracted, re-colored, resized (for high-DPI screens),
' and rendered onto UI elements at run-time.
'
'Menu icons currently lean on the clsMenuImage class by Leandro Ascierto.  Please see that class for details on how
' it works. (A link to Leandro's original project can also be found there.)
'
'This module also handles the rendering of dynamic form, program, and taskbar icons.  When an image is loaded and active,
' those icons can change to match the current image.  For an overview on how this works, visit this MSDN page:
' http://support.microsoft.com/kb/318876
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'API calls for building icons and cursors at run-time
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal cPlanes As Long, ByVal cBitsPerPel As Long, ByVal lpvBits As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateIconIndirect Lib "user32" (icoInfo As ICONINFO) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LoadImageW Lib "user32" (ByVal hInst As Long, ByVal lpsz As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Enum DrawIconEx_Flags
    DI_COMPAT = &H4         'This flag is ignored
    DI_DEFAULTSIZE = &H8    'Draws the icon or cursor using the width and height specified by the system metric values for icons, if the cxWidth and cyWidth parameters are set to zero. If this flag is not specified and cxWidth and cyWidth are set to zero, the function uses the actual resource size.
    DI_IMAGE = &H2          'Draws the icon or cursor using the image.
    DI_MASK = &H1           'Draws the icon or cursor using the mask.
    DI_NOMIRROR = &H10      'Draws the icon as an unmirrored icon. By default, the icon is drawn as a mirrored icon if hdc is mirrored.
    DI_NORMAL = &H3         'Combination of DI_IMAGE and DI_MASK.
End Enum

#If False Then
    Private Const DI_COMPAT = &H4, DI_DEFAULTSIZE = &H8, DI_IMAGE = &H2, DI_MASK = &H1, DI_NOMIRROR = &H10, DI_NORMAL = &H3
#End If

Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As DrawIconEx_Flags) As Long

'System constants for retrieving system default icon sizes and related metrics
Private Const SM_CXICON As Long = 11
Private Const SM_CYICON As Long = 12
Private Const SM_CXCURSOR As Long = 13
Private Const SM_CYCURSOR As Long = 14
Private Const SM_CXSMICON As Long = 49
Private Const SM_CYSMICON As Long = 50
Private Const IMAGE_ICON As Long = 1
Private Const WM_SETICON As Long = &H80
Private Const ICON_SMALL As Long = 0
Private Const ICON_BIG As Long = 1

'Type required to create an icon on-the-fly
Private Type ICONINFO
    fIcon As Boolean
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

'Used to apply and manage custom cursors (without subclassing)
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, ByRef dstInfo As ICONINFO) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

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

'As of v7.0, icon creation and destruction is tracked locally.
Private m_IconsCreated As Long, m_IconsDestroyed As Long

'This constant is used for testing only.  It should always be set to TRUE for production code.
Private Const ALLOW_DYNAMIC_ICONS As Boolean = True

'This array tracks the resource identifiers and consequent numeric identifiers of all loaded icons.  The size of the array
' is arbitrary; just make sure it's larger than the max number of loaded icons.
Private m_IconNames(0 To 511) As String

'We also need to track how many icons have been loaded; this counter will also be used to reference icons in the database
Private m_curIconIndex As Long

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

'PD icon overlay for the task bar icon
Private m_PDIconOverlay As pdDIB

'Load all the menu icons from PhotoDemon's embedded resource file
Public Sub LoadMenuIcons(Optional ByVal alsoApplyMenuIcons As Boolean = True)

    FreeMenuIconCache
    
    With cMenuImage
        
        'Menu icons rely on subclassing in XP (because there's no native OS support for 32-bit menu icons).
        ' If inside the IDE on XP, don't even attempt to place menu icons inside the IDE.
        If (Not OS.IsVistaOrLater) And (Not OS.IsProgramCompiled) Then
            PDDebug.LogAction "XP + IDE detected.  Menu icons are disabled for this session."
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
    m_curIconIndex = 0
    Erase m_IconNames
    
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
    For i = 0 To m_curIconIndex

        If Strings.StringsEqual(m_IconNames(i), resID, False) Then
            iconAlreadyLoaded = True
            iconLocation = i
            Exit For
        End If

    Next i
    
    'If the icon was not found, load it and add it to the list
    If (Not iconAlreadyLoaded) Then
        AddImageResourceToClsMenu resID, cMenuImage
        m_IconNames(m_curIconIndex) = resID
        iconLocation = m_curIconIndex
        m_curIconIndex = m_curIconIndex + 1
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

Private Sub AddImageResourceToClsMenu(ByRef srcResID As String, ByRef targetMenuObject As clsMenuImage, Optional ByVal desiredSize As Long = 0, Optional ByVal desiredPadding As Single = 0.5!)

    'First, attempt to load the image from our internal resource manager
    Dim loadedInternally As Boolean: loadedInternally = False
    If (Not g_Resources Is Nothing) Then
        If g_Resources.AreResourcesAvailable Then
            Dim tmpDIB As pdDIB
            If (desiredSize = 0) Then desiredSize = Interface.FixDPI(16)
            loadedInternally = g_Resources.LoadImageResource(srcResID, tmpDIB, desiredSize, desiredSize, desiredPadding, True, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic))
            If loadedInternally Then targetMenuObject.AddImageFromDIB tmpDIB
        End If
    End If
    
    'If that fails, I probably made a typo in the resource name - note this!
    If (Not loadedInternally) Then PDDebug.LogAction "WARNING!  IconsAndCursors.AddImageResourceToClsMenu failed on: " & srcResID
    
End Sub

'Convert a DIB - any DIB! - to an icon via CreateIconIndirect.  Transparency will be preserved, and by default, the icon will be created
' at the current image's size (though you can specify a custom size if you wish).  Ideally, the passed DIB will have been created using
' the pdImage function "RequestThumbnail".
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
        If (iconSize <= 32) Then resampleMode = GP_IM_HighQualityBicubic Else resampleMode = UserPrefs.GetThumbnailInterpolationPref()
        GDI_Plus.GDIPlus_StretchBlt tmpDIB, 0, 0, iconSize, iconSize, srcDIB, 0, 0, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, , resampleMode, , , , True
        
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

Public Function GetMenuImageCount() As Long
    If (Not cMenuImage Is Nothing) Then GetMenuImageCount = GetMenuImageCount + cMenuImage.ImageCount
    If (Not cMRUIcons Is Nothing) Then GetMenuImageCount = GetMenuImageCount + cMRUIcons.ImageCount
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
            hIcon16 = GetIconFromDIB(thumbDIB, 16)
            
            'Overlay the PD logo on the taskbar icon
            If (m_PDIconOverlay Is Nothing) Then g_Resources.LoadImageResource "pd_icon_glow_16", m_PDIconOverlay, 16, 16
            If (Not m_PDIconOverlay Is Nothing) Then m_PDIconOverlay.AlphaBlendToDC thumbDIB.GetDIBDC, , thumbDIB.GetDIBWidth - 16, thumbDIB.GetDIBHeight - 16
            hIcon32 = GetIconFromDIB(thumbDIB, 32)
            
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

'Given an image in PD's theme file, return it as a system cursor handle.
'PLEASE NOTE: PD auto-caches created cursors, and retains them for the life of the program.  They should not be manually freed.
Private Function CreateCursorFromResource(ByRef resTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0) As Long
    
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
        
        CreateCursorFromResource = CreateCursorFromDIB(resDIB, curHotspotX, curHotspotY)
        
    Else
        PDDebug.LogAction "WARNING!  IconsAndCursors.CreateCursorFromResource failed to find the resource: " & resTitle
    End If
    
End Function

'Given an arbitrary DIB, return a valid cursor handle.  All resources required for creation were auto-freed (except the
' incoming source DIB, obviously), but note that *YOU* are responsible for freeing the cursor handle when finished.
' (This is why this function is not public; PD uses safe wrapper functions that auto-cache and free cursors as relevant.)
Private Function CreateCursorFromDIB(ByRef srcDIB As pdDIB, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0) As Long
    
    'Generate a blank monochrome mask to pass to the icon creation function.
    ' (This is a stupid Windows workaround for 32bpp cursors.  The cursor creation function always assumes
    '  the presence of a mask bitmap, so we have to submit one even if we want the PNG's alpha channel
    '  used for transparency.)
    Dim monoBmp As Long
    monoBmp = CreateBitmap(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 1, 1, ByVal 0&)
    
    'Create an icon header and point it at our temporary mask and original DIB resource
    Dim icoInfo As ICONINFO
    With icoInfo
        .fIcon = False
        .xHotspot = curHotspotX
        .yHotspot = curHotspotY
        .hbmMask = monoBmp
        .hbmColor = srcDIB.GetDIBHandle
    End With
    
    'Create the cursor
    CreateCursorFromDIB = CreateNewIcon(icoInfo)
    
    'Release our temporary mask
    DeleteObject monoBmp
    
End Function

'Retrieve the current system cursor size, in pixels.  Make sure to read the function details - the size returned differs
' from what bare WAPI functions return, by design.
Public Function GetSystemCursorSizeInPx() As Long

    'If this is the first custom cursor request, cache the current system cursor size.
    ' (This function matches those sizes automatically, and the caller does not have control over it, by design.)
    
    'Also, note that Windows cursors typically only use one quadrant of the current system cursor size.  This odd behavior
    ' is why we divide the retrieved cursor size by two.
    If (m_CursorSize = 0) Then
        m_CursorSize = GetSystemMetrics(SM_CYCURSOR) \ 2
        If (m_CursorSize <= 0) Then m_CursorSize = Interface.FixDPI(16)
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

'Set a single form to use the arrow cursor
Public Sub SetArrowCursor(ByRef tControl As Object)
    tControl.MousePointer = vbCustom
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_ARROW)
End Sub

'If a custom 32-bpp cursor has not been loaded, this function will load the resource, convert it to cursor format,
' then store the cursor resource for future reference (so the image doesn't have to be loaded again).
Public Function RequestCustomCursor(ByRef resCursorName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0) As Long

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
        
        Dim tmpHandle As Long, cacheHandle As Boolean
        
        'PD uses special names for some internal cursors.  These are *not* resources, but they are assembled at run-time.
        If Strings.StringsEqual(resCursorName, "HAND-AND-RESIZE", True) Then
            tmpHandle = GetHandAndResizeCursor()
            cacheHandle = True
        Else
            tmpHandle = CreateCursorFromResource(resCursorName, cursorHotspotX, cursorHotspotY)
            cacheHandle = (tmpHandle <> 0)
        End If
        
        If (tmpHandle <> 0) And cacheHandle Then
            ReDim Preserve m_customCursorNames(0 To m_numCustomCursors) As String
            ReDim Preserve m_customCursorHandles(0 To m_numCustomCursors) As Long
            m_customCursorNames(m_numCustomCursors) = resCursorName
            m_customCursorHandles(m_numCustomCursors) = tmpHandle
            m_numCustomCursors = m_numCustomCursors + 1
        End If
            
        RequestCustomCursor = tmpHandle
        
    End If

End Function

'Given an image in the .exe's resource section (typically a 32-bpp image), load it to a pdDIB object.
Public Function LoadResourceToDIB(ByRef resTitle As String, ByRef dstDIB As pdDIB, Optional ByVal desiredWidth As Long = 0, Optional ByVal desiredHeight As Long = 0, Optional ByVal desiredBorders As Long = 0, Optional ByVal useCustomColor As Long = -1, Optional ByVal suspendMonochrome As Boolean = False, Optional ByVal resampleAlgorithm As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal usePDResamplerInstead As PD_ResamplingFilter = rf_Automatic) As Boolean
    
    'Some functions may call this before GDI+ has loaded; exit if that happens
    If Drawing2D.IsRenderingEngineActive(P2_GDIPlusBackend) Then
    
        'Make sure PD's resource manager is also active before attempting the load
        If (Not g_Resources Is Nothing) Then
            If g_Resources.AreResourcesAvailable Then
                LoadResourceToDIB = g_Resources.LoadImageResource(resTitle, dstDIB, desiredWidth, desiredHeight, desiredBorders, , useCustomColor, suspendMonochrome, resampleAlgorithm, usePDResamplerInstead)
            End If
        End If
        
        'If we failed to find the requested resource, return a small blank DIB (to work around some old code
        ' that expects an initialized DIB), report the problem, then exit.  (As of v7.0, this failsafe check
        ' will only be triggered by misspelling a resource name!)
        If (Not LoadResourceToDIB) Then
            If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
            dstDIB.CreateBlank 16, 16, 32, 0, 0
            LoadResourceToDIB = False
            PDDebug.LogAction "WARNING!  LoadResourceToDIB couldn't find <" & resTitle & ">.  Check your spelling and try again."
        End If
        
    Else
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
        dstLargeIcon = SendMessageA(targetHWnd, WM_SETICON, ICON_BIG, ByVal hIconLarge)
        dstSmallIcon = SendMessageA(targetHWnd, WM_SETICON, ICON_SMALL, ByVal hIconSmall)
    End If
End Sub

Public Sub MirrorCurrentIconsToWindow(ByVal targetHWnd As Long, Optional ByVal setLargeIconOnly As Boolean = False, Optional ByRef dstSmallIcon As Long = 0, Optional ByRef dstLargeIcon As Long = 0)
    If PDImages.IsImageActive() Then
        ChangeWindowIcon targetHWnd, IIf(setLargeIconOnly, 0&, PDImages.GetActiveImage.GetImageIcon(False)), PDImages.GetActiveImage.GetImageIcon(True), dstSmallIcon, dstLargeIcon
    Else
        ChangeWindowIcon targetHWnd, IIf(setLargeIconOnly, 0&, m_DefaultIconSmall), m_DefaultIconLarge, dstSmallIcon, dstLargeIcon
    End If
End Sub

'When all images are unloaded (or when the program is first loaded), we must reset the program icon to its default values.
Public Sub ResetAppIcons()
    
    Const DEFAULT_ICON As String = "AAA"
    
    If (m_DefaultIconLarge = 0) Then
        m_DefaultIconLarge = LoadImageW(App.hInstance, StrPtr(DEFAULT_ICON), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), 0&)
    End If
    
    If (m_DefaultIconSmall = 0) Then
        m_DefaultIconSmall = LoadImageW(App.hInstance, StrPtr(DEFAULT_ICON), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0&)
    End If
    
    ChangeAppIcons m_DefaultIconSmall, m_DefaultIconLarge
    
End Sub

'When PD is first loaded, we associate an icon with the hidden "ThunderMain" owner window,
' to ensure proper icons in places like Task Manager.
Public Sub SetThunderMainIcon()

    'Start by loading the default icons from the resource file, as necessary
    ResetAppIcons
    
    Dim tmHWnd As Long
    tmHWnd = OS.ThunderMainHWnd()
    SendMessageA tmHWnd, WM_SETICON, ICON_SMALL, ByVal m_DefaultIconSmall
    SendMessageA tmHWnd, WM_SETICON, ICON_BIG, ByVal m_DefaultIconLarge

End Sub

Private Function GetSysCursorAsDIB(ByVal cursorType As SystemCursorConstant, ByRef dstDIB As pdDIB, ByRef dstBounds As RectF, Optional ByVal initDIBForMe As Boolean = True, Optional ByVal renderFrame As Long = 0) As Boolean

    'First, retrieve a handle to the system cursor in question
    Dim hCursor As Long
    hCursor = LoadCursor(0&, cursorType)
    
    'Next, prepare the destination DIB
    If initDIBForMe Or (dstDIB Is Nothing) Then
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank GetSystemMetrics(SM_CXCURSOR), GetSystemMetrics(SM_CYCURSOR), 32, 0, 0
        dstDIB.SetInitialAlphaPremultiplicationState False
    End If
    
    'Use DrawIconEx to render the cursor into the 32-bpp DIB, then premultiply the result
    GetSysCursorAsDIB = (DrawIconEx(dstDIB.GetDIBDC, 0, 0, hCursor, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, renderFrame, 0, DI_NORMAL) <> 0)
    dstDIB.SetAlphaPremultiplication True
    If (Not GetSysCursorAsDIB) Then PDDebug.LogAction "WARNING!  IconsAndCursors.GetSysCursorAsDIB failed on DrawIconEx."
    
    'We now want to see if the destination DIB contains valid data.  (We define this as meeting two criteria:
    ' 1) the image must contain at least some non-transparent bytes, and...
    ' 2) the image must not be one solid uniformly colored block.  (This is a telltale sign of GDI failing in 32-bpp mode.)
    ' If the image does not appear to be valid, we will manually assemble it using legacy mask+bmp merging.
    If DIBs.IsDIBAlphaBinary(dstDIB, False) Or DIBs.IsDIBSolidColor(dstDIB) Then
        
        'This DIB does not contain a usable alpha channel, so we need to render it manually in two stages.  Start by creating two
        ' temporary DIBs - and importantly, make sure the DIBs are *24-bpp*.
        Dim tmpMask As pdDIB
        Set tmpMask = New pdDIB
        tmpMask.CreateBlank GetSystemMetrics(SM_CXCURSOR), GetSystemMetrics(SM_CYCURSOR), 24, 0
        
        Dim tmpBMP As pdDIB
        Set tmpBMP = New pdDIB
        tmpBMP.CreateBlank GetSystemMetrics(SM_CXCURSOR), GetSystemMetrics(SM_CYCURSOR), 24, 0
        
        'Render the cursor mask into the mask DIB, and the cursor image bytes into the BMP DIB
        GetSysCursorAsDIB = (DrawIconEx(tmpMask.GetDIBDC, 0, 0, hCursor, tmpMask.GetDIBWidth, tmpMask.GetDIBHeight, renderFrame, 0, DI_MASK) <> 0)
        If GetSysCursorAsDIB Then GetSysCursorAsDIB = (DrawIconEx(tmpBMP.GetDIBDC, 0, 0, hCursor, tmpBMP.GetDIBWidth, tmpBMP.GetDIBHeight, renderFrame, 0, DI_IMAGE) <> 0)
        
        If GetSysCursorAsDIB Then
            
            'We now want to merge the two 24-bpp DIBs into a usable 32-bpp DIB.
            
            'Start by copying a grayscale copy of the DIB mask into a byte array.  We will use this data to produce the
            ' 32-bpp merged image's alpha channel.
            Dim transBytes() As Byte
            If DIBs.GetDIBGrayscaleMap(tmpMask, transBytes, False) Then
            
                'Note that the mask will be inverted, by default.  (Black is opaque and white is transparent.)
                ' We need to reverse it.
                Filters_ByteArray.InvertByteArray transBytes, tmpMask.GetDIBWidth, tmpMask.GetDIBHeight
            
                'Next, copy the image bytes into the destination DIB we created in the first place.
                tmpBMP.ConvertTo32bpp
                dstDIB.CreateFromExistingDIB tmpBMP
                
                'Finally, merge the mask as a transparency channel, and premultiply the end result
                DIBs.ApplyTransparencyTable dstDIB, transBytes
                dstDIB.SetAlphaPremultiplication True, True
                
                GetSysCursorAsDIB = True
            
            End If
        
        Else
            PDDebug.LogAction "WARNING!  IconsAndCursors.GetSysCursorAsDIB failed on DrawIconEx, second attempt(s)."
        End If
        
    End If
    
    'One way or another, we've hopefully ended up with a usable 32-bpp cursor by now.
    If GetSysCursorAsDIB Then
    
        'The last thing we want to do is determine the usable area of the DIB consumed by the cursor.
        ' (Some system cursors, such as the default Windows arrow, may occupy less than 1/4 of the default
        ' system cursor size.)
        GetSysCursorAsDIB = DIBs.GetRectOfInterest(dstDIB, dstBounds)
    
    End If
    
End Function

'Some PD objects are clickable *and* draggable; they use a specialized "hand+resize" cursor that we generate on-the-fly
' from the current hand and resize system cursors.
'
'NOTE!  As of 28 February 2018, this function is considered "disabled".  There are persistent and difficult to predict
' issues with rendering system cursors into custom containers, particularly when anything but the non-default system
' cursor theme is in use.  I need to do much more investigation before enabling this function.  (See pdTitleBar's
' _MouseEnter event for a location where this could potential be useful, if the various kinks are worked out.)
Private Function GetHandAndResizeCursor() As Long

    'Start by retrieving the two cursors in question as DIBs
    Dim handDIB As pdDIB, handRect As RectF
    Dim resizeDIB As pdDIB, resizeRect As RectF
    GetHandAndResizeCursor = GetSysCursorAsDIB(IDC_HAND, handDIB, handRect)
    If GetHandAndResizeCursor Then GetHandAndResizeCursor = GetSysCursorAsDIB(IDC_SIZENS, resizeDIB, resizeRect)
    
    'We now need to retrieve the cursor hotspot for the current handDIB, because we'll be reusing that for our cursor
    Dim handInfo As ICONINFO
    If GetHandAndResizeCursor Then GetHandAndResizeCursor = (GetIconInfo(LoadCursor(0&, IDC_HAND), handInfo) <> 0)
    If GetHandAndResizeCursor Then
    
        'We now have everything we need to assemble a new icon!  Start by combining the two source DIBs into a new
        ' destination DIB.
        Dim newDIB As pdDIB
        Set newDIB = New pdDIB
        newDIB.CreateBlank GetSystemMetrics(SM_CXCURSOR) * 2, GetSystemMetrics(SM_CYCURSOR) * 2, 32, 0, 0
        newDIB.SetInitialAlphaPremultiplicationState True
        GDI.BitBltWrapper newDIB.GetDIBDC, 0, 0, handDIB.GetDIBWidth, handDIB.GetDIBHeight, handDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'The arrow DIB is a little different; because the arrow cursor can be vastly different sizes (depending on the
        ' current system cursor theme), we want to make sure we only grab the *relevant* portion of the arrow DIB,
        ' using the boundary rect returned by GetSysCursorAsDIB.
        Const arrowShrink As Double = 0.25
        Dim dibWidth As Long, dibHeight As Long, dibPosX As Long, dibPosY As Long
        dibPosX = handRect.Left + handRect.Width + Interface.FixDPI(3)
        dibPosY = handRect.Top + handRect.Height * arrowShrink
        dibHeight = handRect.Height * (1# - arrowShrink)
        dibWidth = (resizeRect.Width / resizeRect.Height) * dibHeight
        
        With resizeRect
            GDI_Plus.GDIPlus_StretchBlt newDIB, dibPosX, dibPosY, dibWidth, dibHeight, resizeDIB, .Left, .Top, .Width, .Height, , GP_IM_HighQualityBicubic, , True
        End With
        
        'Make a cursor from the composited DIB
        GetHandAndResizeCursor = CreateCursorFromDIB(newDIB, handInfo.xHotspot, handInfo.yHotspot)
        
        'Make sure no ddbs are leaked; GetIconInfo allocates two of 'em!
        If (handInfo.hbmColor <> 0) Then DeleteObject handInfo.hbmColor
        If (handInfo.hbmMask <> 0) Then DeleteObject handInfo.hbmMask
        
    Else
        PDDebug.LogAction "WARNING!  IconsAndCursors.GetHandAndResizeCursor failed."
    End If

End Function
