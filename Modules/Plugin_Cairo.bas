Attribute VB_Name = "Plugin_Cairo"
'***************************************************************************
'Cairo library interface
'Copyright 2018-2018 by Tanner Helland
'Created: 21/June/18
'Last updated: 25/June/18
'Last update: continued work on initial build
'
'While PhotoDemon provides manual implementations of just about every required graphics op in the program,
' it is sometimes much faster (and/or easier) to lean on 3rd-party libraries.  Cairo in particular has
' excellent support for masking - a feature that GDI+ lacks, which is an unfortunate headache for us.
'
'As part of the 7.2 release, I've started custom-building cairo and shipping it alongside PD as an
' optional implementation for certain features.  Because Cairo itself is LGPL/MPL-licensed, I'm not
' making any special changes to the library - just compiling it as stdcall with name-mangling resolved.
' At present, any version of the library from the past decade or so should work, provided it meets those
' criteria.  Feel free to drop in your own version of the library, or to drop in any other stdcall-based
' wrapper, like Olaf Schmidt's popular version at http://www.vbrichclient.com/#/en/Downloads.htm
' (but note that you'll need to either rename his DLL, or rename this module's function declares to
' "vb_cairo_sqlite.dll" for his version to work).
'
'I've had trouble with cairo crashing on XP (but not Win 7 on identical hardware), and rather than spend
' an inordinate amount of time debugging the problem, I've simply disabled cairo integration on XP and Vista.
' If you encounter problems with the wrapper on a newer version of Windows, please let me know and I'll
' investigate further.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Pixel formats for cairo surfaces.  Note that PD exclusively uses ARGB32 surfaces; other surface formats
' are *not* currently well-tested.
Private Enum Cairo_Format
    cf_Invalid = -1
    cf_ARGB32 = 0
    cf_RGB24 = 1
    cf_A8 = 2
    cf_A1 = 3
    cf_RGB16_565 = 4
    cf_RGB30 = 5
End Enum

#If False Then
    Private Const cf_Invalid = -1, cf_ARGB32 = 0, cf_RGB24 = 1, cf_A8 = 2, cf_A1 = 3, cf_RGB16_565 = 4, cf_RGB30 = 5
#End If

'Cairo blend operators
Private Enum Cairo_Operator
    co_Clear = 0
    co_Source = 1
    co_Over = 2
    co_In = 3
    co_Out = 4
    co_Atop = 5
    co_Dest = 6
    co_DestOver = 7
    co_DestIn = 8
    co_DestOut = 9
    co_DestAtop = 10
    co_XOR = 11
    co_Add = 12
    co_Saturate = 13
    co_Multiply = 14
    co_Screen = 15
    co_Overlay = 16
    co_Darken = 17
    co_Lighten = 18
    co_ColorDodge = 19
    co_ColorBurn = 20
    co_HardLight = 21
    co_SoftLight = 22
    co_Difference = 23
    co_Exclusion = 24
    co_HSLHue = 25
    co_HSLSaturation = 26
    co_HSLColor = 27
    co_HSLLuminosity = 28
End Enum

#If False Then
    Private Const co_Clear = 0, co_Source = 1, co_Over = 2, co_In = 3, co_Out = 4, co_Atop = 5, co_Dest = 6, co_DestOver = 7, co_DestIn = 8, co_DestOut = 9
    Private Const co_DestAtop = 10, co_XOR = 11, co_Add = 12, co_Saturate = 13, co_Multiply = 14, co_Screen = 15, co_Overlay = 16, co_Darken = 17, co_Lighten = 18, co_ColorDodge = 19
    Private Const co_ColorBurn = 20, co_HardLight = 21, co_SoftLight = 22, co_Difference = 23, co_Exclusion = 24, co_HSLHue = 25, co_HSLSaturation = 26, co_HSLColor = 27, co_HSLLuminosity = 28
#End If

Public Enum Cairo_Filter
    cf_Fast = 0
    cf_Good = 1
    cf_Best = 2
    cf_Nearest = 4
    cf_Bilinear = 5
End Enum

#If False Then
    Private Const cf_Fast = 0, cf_Good = 1, cf_Best = 2, cf_Nearest = 4, cf_Bilinear = 5
#End If

'Exported cairo functions
Private Declare Sub cairo_clip_extents Lib "cairo" (ByVal dstContext As Long, ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double)
Private Declare Function cairo_create Lib "cairo" (ByVal dstSurface As Long) As Long
Private Declare Sub cairo_destroy Lib "cairo" (ByVal srcContext As Long)
Private Declare Function cairo_image_surface_create_for_data Lib "cairo" (ByVal ptrToPixels As Long, ByVal pxFormat As Cairo_Format, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal imgStride As Long) As Long
Private Declare Sub cairo_paint Lib "cairo" (ByVal dstContext As Long)
Private Declare Function cairo_pattern_create_for_surface Lib "cairo" (ByVal srcSurface As Long) As Long
Private Declare Sub cairo_pattern_destroy Lib "cairo" (ByVal srcPattern As Long)
Private Declare Sub cairo_pattern_set_filter Lib "cairo" (ByVal dstPattern As Long, ByVal newFilter As Cairo_Filter)
Private Declare Sub cairo_scale Lib "cairo" (ByVal dstContext As Long, ByVal scaleX As Double, ByVal scaleY As Double)
Private Declare Sub cairo_set_operator Lib "cairo" (ByVal dstContext As Long, ByVal newOperator As Cairo_Operator)
Private Declare Sub cairo_set_source Lib "cairo" (ByVal dstContext As Long, ByVal srcPattern As Long)
Private Declare Sub cairo_set_source_rgba Lib "cairo" (ByVal dstContext As Long, ByVal srcRed As Double, ByVal srcGreen As Double, ByVal srcBlue As Double, ByVal srcAlpha As Double)
Private Declare Sub cairo_set_source_surface Lib "cairo" (ByVal dstContext As Long, ByVal srcSurface As Long, ByVal patternOriginX As Double, ByVal patternOriginY As Double)
Private Declare Sub cairo_surface_destroy Lib "cairo" (ByVal srcSurface As Long)
Private Declare Sub cairo_surface_set_device_offset Lib "cairo" (ByVal dstSurface As Long, ByVal xOffset As Double, ByVal yOffset As Double)
Private Declare Sub cairo_translate Lib "cairo" (ByVal dstContext As Long, ByVal transX As Double, ByVal transY As Double)
Private Declare Function cairo_version_string Lib "cairo" () As Long
Private Declare Function cairo_win32_surface_create Lib "cairo" (ByVal dstDC As Long) As Long

'Persistent LoadLibrary handle; will be non-zero if cairo has been loaded.
Private m_hLibCairo As Long

'Initialize Cairo.  Do not call this until you have verified the dll's existence (typically via the PluginManager module)
Public Function InitializeCairo() As Boolean
    
    'Due to current null-pointer crashes on XP (which I have tried and failed to resolve),
    ' Cairo support is limited to Win 7+.  It's possible that the library will also run fine on Vista,
    ' but without an active test rig, I'm not going to risk it.
    If OS.IsWin7OrLater Then
        
        If (m_hLibCairo = 0) Then
        
            'Manually load the DLL from the plugin folder (should be App.Path\App\PhotoDemon\Plugins)
            Dim cairoPath As String
            cairoPath = PluginManager.GetPluginPath & "cairo.dll"
            m_hLibCairo = VBHacks.LoadLib(cairoPath)
            InitializeCairo = (m_hLibCairo <> 0)
            
            If (Not InitializeCairo) Then
                PDDebug.LogAction "WARNING!  LoadLibrary failed to load cairo.  Last DLL error: " & Err.LastDllError
                PDDebug.LogAction "(FYI, the attempted path was: " & cairoPath & ")"
            End If
            
        Else
            InitializeCairo = True
        End If
        
    Else
        InitializeCairo = False
    End If
    
End Function

'When PD closes, be a good citizen and release our library handle!
Public Sub ReleaseCairo()
    If (m_hLibCairo <> 0) Then VBHacks.FreeLib m_hLibCairo
End Sub

Public Function GetCairoVersion() As String
    If (m_hLibCairo <> 0) Then GetCairoVersion = Strings.StringFromCharPtr(cairo_version_string(), False) Else GetCairoVersion = g_Language.TranslateMessage("this plugin is not compatible with your version of Windows")
End Function

'Cairo-based StretchBlt.  IMPORTANTLY, this function does not work if the source and destination
' DIBs are identical - the intermediary results of the Blt will be copied as the function proceeds!
' I don't currently know an easy workaround for this.
Public Sub Cairo_StretchBlt(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcDIB As pdDIB, ByVal x2 As Single, ByVal y2 As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal filterType As Cairo_Filter = cf_Best, Optional ByVal useThisDestinationDCInstead As Long = 0, Optional ByVal disableEdgeFix As Boolean = False, Optional ByVal isZoomedIn As Boolean = False, Optional ByVal dstCopyIsOkay As Boolean = False)
    
    If (dstDIB Is Nothing) And (useThisDestinationDCInstead = 0) Then Exit Sub
    
    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Currency
    'VBHacks.GetHighResTime profileTime
    
    'Create a cairo surface object that points to the destination DIB's DC
    Dim dstSurface As Long, dstContext As Long
    If (useThisDestinationDCInstead <> 0) Then
        dstSurface = Plugin_Cairo.WrapCairoSurfaceAroundDC(dstSurface, dstContext)
    Else
        dstSurface = Plugin_Cairo.GetCairoSurfaceFromPDDib(dstDIB, dstContext)
    End If
    
    'Set the offset for the destination
    cairo_surface_set_device_offset dstSurface, x1, y1
    
    'Set the scaling factor for the transform
    cairo_scale dstContext, dstWidth / srcWidth, dstHeight / srcHeight
    
    'If copying is okay, set the context blend accordingly
    If dstCopyIsOkay Then cairo_set_operator dstContext, co_Source
    
    'Next, we need a pattern that points at the source image.  Note that we apply the source offset now.
    Dim srcSurface As Long, srcPattern As Long
    srcSurface = GetCairoSurfaceFromPDDib_NoContext(srcDIB)
    cairo_surface_set_device_offset srcSurface, x2, y2
    srcPattern = Pattern_GetFromSurface(srcSurface)
    
    'Request the resize filter we were passed
    Pattern_SetResampleFilter srcPattern, filterType
    
    'Set the pattern; note that this locks-in all current transform settings
    cairo_set_source dstContext, srcPattern
    
    'Paint the pattern
    cairo_paint dstContext
    
    'Delete everything
    cairo_pattern_destroy srcPattern
    cairo_surface_destroy srcSurface
    cairo_surface_destroy dstSurface
    cairo_destroy dstContext
    
'    'We now need to create a transform that describes the StretchBlt parameters
'
'        'To fix antialiased fringing around image edges, specify a wrap mode.  This will prevent the faulty GDI+ resize
'        ' algorithm from drawing semi-transparent lines randomly around image borders.
'        ' Thank you to http://stackoverflow.com/questions/1890605/ghost-borders-ringing-when-resizing-in-gdi for the fix.
'        Dim imgAttributesHandle As Long
'        GdipCreateImageAttributes imgAttributesHandle
'
'        'To improve performance, explicitly request high-speed (aka linear) alpha compositing operation, and standard
'        ' pixel offsets (on pixel borders, instead of center points)
'        If (Not disableEdgeFix) Then GdipSetImageAttributesWrapMode imgAttributesHandle, GP_WM_TileFlipXY, 0, 0
'        GdipSetCompositingQuality hGraphics, GP_CQ_AssumeLinear
'        If isZoomedIn Then GdipSetPixelOffsetMode hGraphics, GP_POM_HighQuality Else GdipSetPixelOffsetMode hGraphics, GP_POM_HighSpeed
'
'        'If modified alpha is requested, pass the new value to this image container
'        If (newAlpha < 1!) Then
'            m_AttributesMatrix(3, 3) = newAlpha
'            GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1, VarPtr(m_AttributesMatrix(0, 0)), 0, GP_CMF_Default
'        End If
'
'        'If the caller doesn't care about source blending (e.g. they're painting to a known transparent destination),
'        ' copy mode can improve performance.
'        If dstCopyIsOkay Then GdipSetCompositingMode hGraphics, GP_CM_SourceCopy
'
'        'Because the resize step is the most cumbersome one, it can be helpful to track it
'        'Dim resizeTime As Currency
'        'VBHacks.GetHighResTime resizeTime
'
'        'Perform the resize
'        GdipDrawImageRectRect hGraphics, hBitmap, x1, y1, dstWidth, dstHeight, x2, y2, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&
'
'        'Report resize time here
'        'Debug.Print "GDI+ resize time: " & Format(CStr(VBHacks.GetTimerDifferenceNow(resizeTime) * 1000), "0000.00") & " ms"
'
'        'Release our image attributes object
'        GdipDisposeImageAttributes imgAttributesHandle
'
'        'Reset alpha in the master identity matrix
'        If (newAlpha < 1!) Then m_AttributesMatrix(3, 3) = 1!
'
'        'Update premultiplication status in the target
'        If (Not dstDIB Is Nothing) Then dstDIB.SetInitialAlphaPremultiplicationState srcDIB.GetAlphaPremultiplication
'
'    End If
'
'    'Release both the destination graphics object and the source bitmap object
'    GdipDisposeImage hBitmap
'    GdipDeleteGraphics hGraphics
    
    'To keep resources low, free the destination DIB from its DC
    If (Not dstDIB Is Nothing) Then dstDIB.FreeFromDC
    
    'Uncomment the line below to receive timing reports
    'Debug.Print "GDI+ wrapper time: " & Format(CStr(VBHacks.GetTimerDifferenceNow(profileTime) * 1000), "0000.00") & " ms"
    
End Sub

'Only works on 32-bpp DIBs at present
Public Function GetCairoSurfaceFromPDDib(ByRef srcDIB As pdDIB, ByRef dstContext As Long) As Long
    If (srcDIB.GetDIBColorDepth = 32) Then
        GetCairoSurfaceFromPDDib = cairo_image_surface_create_for_data(srcDIB.GetDIBPointer, cf_ARGB32, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBStride)
        If (GetCairoSurfaceFromPDDib <> 0) Then dstContext = cairo_create(GetCairoSurfaceFromPDDib)
    Else
        GetCairoSurfaceFromPDDib = 0
    End If
End Function

'Only works on 32-bpp DIBs at present
Public Function GetCairoSurfaceFromPDDib_NoContext(ByRef srcDIB As pdDIB) As Long
    If (srcDIB.GetDIBColorDepth = 32) Then
        GetCairoSurfaceFromPDDib_NoContext = cairo_image_surface_create_for_data(srcDIB.GetDIBPointer, cf_ARGB32, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBStride)
    Else
        GetCairoSurfaceFromPDDib_NoContext = 0
    End If
End Function

'Wrap a cairo surface around an arbitrary DC.  Note that the destination DC will *always*
' be treated as 24-bpp, due to legacy GDI conventions - so do *not* use this to wrap pdDIB objects.
' Use GetCairoSurfaceFromPDDib, instead.
Public Function WrapCairoSurfaceAroundDC(ByVal dstDC As Long, ByRef dstContext As Long) As Long
    WrapCairoSurfaceAroundDC = cairo_win32_surface_create(dstDC)
    If (WrapCairoSurfaceAroundDC <> 0) Then dstContext = cairo_create(WrapCairoSurfaceAroundDC)
End Function

'Return a pattern handle to a cairo surface; this pattern can subsequently be used for painting,
' including tasks like resizing.
Public Function Pattern_GetFromSurface(ByVal srcSurface As Long) As Long
    Pattern_GetFromSurface = cairo_pattern_create_for_surface(srcSurface)
End Function

Public Sub Pattern_SetResampleFilter(ByVal dstPattern As Long, ByVal srcFilter As Cairo_Filter)
    cairo_pattern_set_filter dstPattern, srcFilter
End Sub

Public Sub FreeCairoContext(ByRef srcContext As Long)
    'PDDebug.LogAction "cairo_destroy: " & srcContext
    If (srcContext <> 0) Then cairo_destroy srcContext
    srcContext = 0
End Sub

Public Sub FreeCairoPattern(ByRef srcPattern As Long)
    'PDDebug.LogAction "cairo_pattern_destroy: " & srcPattern
    If (srcPattern <> 0) Then cairo_pattern_destroy srcPattern
    srcPattern = 0
End Sub

Public Sub FreeCairoSurface(ByRef srcSurface As Long)
    'PDDebug.LogAction "cairo_surface_destroy: " & srcSurface
    If (srcSurface <> 0) Then cairo_surface_destroy srcSurface
    srcSurface = 0
End Sub

Public Sub TestPainting(ByRef srcDIB As pdDIB)
    
    Dim dstSurface As pd2DSurfaceCairo
    Set dstSurface = New pd2DSurfaceCairo
    If dstSurface.WrapAroundPDDIB(srcDIB) Then
        
        'Set rendering source
        
        'Test plain color rendering:
        'PDDebug.LogAction "cairo_set_source_rgba"
        'cairo_set_source_rgba dstsurface.GetCairoContextHandle(), 0#, 0#, 1#, 0.5
        
        'Test full-image rendering:
        PDDebug.LogAction "cairo_set_source_surface"
        cairo_set_source_surface dstSurface.GetCairoContextHandle(), dstSurface.GetCairoSurfaceHandle(), 0#, 0#
        cairo_set_operator dstSurface.GetCairoContextHandle(), co_Multiply
        
        'Paint
        PDDebug.LogAction "cairo_paint"
        cairo_paint dstSurface.GetCairoContextHandle()
        
        'Free resources
        Set dstSurface = Nothing
        
    End If
    
End Sub

Public Sub TestOnActiveImage()

    If (g_OpenImageCount > 0) Then
        'Plugin_Cairo.TestPainting pdImages(g_CurrentImage).GetActiveDIB
        Dim x1 As Single, y1 As Single, dstWidth As Single, dstHeight As Single
        Dim x2 As Single, y2 As Single, srcWidth As Single, srcHeight As Single
        With pdImages(g_CurrentImage).GetActiveDIB
            x1 = .GetDIBWidth * 0.25
            y1 = .GetDIBHeight * 0.25
            dstWidth = .GetDIBWidth * 0.5
            dstHeight = .GetDIBHeight * 0.5
            x2 = 0
            y2 = 0
            srcWidth = .GetDIBWidth
            srcHeight = .GetDIBHeight
        End With
        
        Plugin_Cairo.Cairo_StretchBlt pdImages(g_CurrentImage).GetActiveDIB, x1, y1, dstWidth, dstHeight, pdImages(g_CurrentImage).GetActiveDIB, x2, y2, srcWidth, srcHeight
        
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_Layer, 0
        ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.MainCanvas(0)
    End If
        
End Sub
