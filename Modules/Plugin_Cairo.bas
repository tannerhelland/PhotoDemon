Attribute VB_Name = "Plugin_Cairo"
'***************************************************************************
'Cairo library interface
'Copyright 2018-2019 by Tanner Helland
'Created: 21/June/18
'Last updated: 03/January/19
'Last update: start serious expansion for required gradient features
'
'While PhotoDemon provides manual implementations of just about every required graphics op in the program,
' it is sometimes much faster (and/or easier) to lean on 3rd-party libraries.  Cairo in particular has
' excellent support for masking - a feature that GDI+ lacks, which is an unfortunate headache for us.
'
'As part of the 7.2 release, I've started shipping a community-built stdcall cairo library with PD.
' (https://github.com/VBForumsCommunity/VbCairo).  Because Cairo itself is LGPL/MPL-licensed, no special
' changes have been made to the library - it is simply compiled as stdcall with name-mangling resolved.
' At present, any version of the library from the past decade or so should work, provided it meets those
' criteria.  Feel free to drop in your own version of the library, or to drop in any other stdcall-based
' wrapper, like Olaf Schmidt's popular version at http://www.vbrichclient.com/#/en/Downloads.htm
' (but note that you'll need to either rename his DLL, or rename this module's function declares to
' "vb_cairo_sqlite.dll" for his version to work).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Antialias behavior; note that subpixel shading does not work for anything but explicit text objects
Public Enum Cairo_Antialias
    ca_DEFAULT = 0
    ca_NONE = 1
    ca_GRAY = 2
    ca_SUBPIXEL = 3
    ca_FAST = 4
    ca_GOOD = 5
    ca_BEST = 6
End Enum

#If False Then
    Private Const ca_DEFAULT = 0, ca_NONE = 1, ca_GRAY = 2, ca_SUBPIXEL = 3, ca_FAST = 4, ca_GOOD = 5, ca_BEST = 6
#End If

'Wrap behavior for things like gradient patterns
Public Enum Cairo_Extend
    ce_ExtendNone = 0
    ce_ExtendRepeat = 1
    ce_ExtendReflect = 2
    ce_ExtendPad = 3
End Enum

#If False Then
    Private Const ce_ExtendNone = 0, ce_ExtendRepeat = 1, ce_ExtendReflect = 2, ce_ExtendPad = 3
#End If

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
Public Enum Cairo_Operator
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
Private Declare Function cairo_create Lib "cairo" (ByVal dstSurface As Long) As Long
Private Declare Sub cairo_destroy Lib "cairo" (ByVal srcContext As Long)

Private Declare Sub cairo_clip_extents Lib "cairo" (ByVal dstContext As Long, ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double)
Private Declare Sub cairo_fill Lib "cairo" (ByVal dstContext As Long)
Private Declare Sub cairo_fill_preserve Lib "cairo" (ByVal dstContext As Long)
Private Declare Function cairo_image_surface_create_for_data Lib "cairo" (ByVal ptrToPixels As Long, ByVal pxFormat As Cairo_Format, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal imgStride As Long) As Long
Private Declare Sub cairo_paint Lib "cairo" (ByVal dstContext As Long)
Private Declare Sub cairo_pattern_add_color_stop_rgb Lib "cairo" (ByVal dstPattern As Long, ByVal srcOffset As Double, ByVal srcRed As Double, ByVal srcGreen As Double, ByVal srcBlue As Double)
Private Declare Sub cairo_pattern_add_color_stop_rgba Lib "cairo" (ByVal dstPattern As Long, ByVal srcOffset As Double, ByVal srcRed As Double, ByVal srcGreen As Double, ByVal srcBlue As Double, ByVal srcAlpha As Double)
Private Declare Function cairo_pattern_create_for_surface Lib "cairo" (ByVal srcSurface As Long) As Long
Private Declare Function cairo_pattern_create_linear Lib "cairo" (ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double) As Long
Private Declare Function cairo_pattern_create_radial Lib "cairo" (ByVal cx0 As Double, ByVal cy0 As Double, ByVal radius0 As Double, ByVal cx1 As Double, ByVal cy1 As Double, ByVal radius1 As Double) As Long
Private Declare Sub cairo_pattern_destroy Lib "cairo" (ByVal srcPattern As Long)
Private Declare Sub cairo_pattern_set_extend Lib "cairo" (ByVal dstPattern As Long, ByVal newExtend As Cairo_Extend)
Private Declare Sub cairo_pattern_set_filter Lib "cairo" (ByVal dstPattern As Long, ByVal newFilter As Cairo_Filter)
Private Declare Sub cairo_rectangle Lib "cairo" (ByVal dstContext As Long, ByVal rX As Double, ByVal rY As Double, ByVal rWidth As Double, ByVal rHeight As Double)
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
Public Sub Cairo_StretchBlt(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcDIB As pdDIB, ByVal x2 As Single, ByVal y2 As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal filterType As Cairo_Filter = cf_Good, Optional ByVal useThisDestinationDCInstead As Long = 0, Optional ByVal disableEdgeFix As Boolean = False, Optional ByVal isZoomedIn As Boolean = False, Optional ByVal dstCopyIsOkay As Boolean = False)
    
    If (dstDIB Is Nothing) And (useThisDestinationDCInstead = 0) Then Exit Sub
    
    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    Dim profileTime As Currency, lastTime As Currency
    VBHacks.GetHighResTime profileTime
    
    'Create a cairo surface object that points to the destination DIB's DC
    Dim dstSurface As Long, dstContext As Long
    If (useThisDestinationDCInstead <> 0) Then
        dstSurface = Plugin_Cairo.WrapCairoSurfaceAroundDC(dstSurface, dstContext)
    Else
        dstSurface = Plugin_Cairo.GetCairoSurfaceFromPDDib(dstDIB, dstContext)
    End If
    
    'Debug.Print "Time required for surface creation: " & VBHacks.GetTimeDiffNowAsString(profileTime)
    'VBHacks.GetHighResTime lastTime
    
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
    
    'Debug.Print "Time required for pattern assembly: " & VBHacks.GetTimeDiffNowAsString(lastTime)
    'VBHacks.GetHighResTime lastTime
    
    'Paint the pattern
    cairo_paint dstContext
    
    'Debug.Print "Time required for paint: " & VBHacks.GetTimeDiffNowAsString(lastTime)
    
    'Delete everything
    cairo_pattern_destroy srcPattern
    cairo_surface_destroy srcSurface
    cairo_surface_destroy dstSurface
    cairo_destroy dstContext
    
    'To keep resources low, free the destination DIB from its DC
    If (Not dstDIB Is Nothing) Then dstDIB.FreeFromDC
    
    'Uncomment the line below to receive timing reports
    'Debug.Print "Total cairo wrapper time: " & VBHacks.GetTimeDiffNowAsString(profileTime)
    
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

Public Function Context_Fill(ByVal dstContext As Long)
    cairo_fill dstContext
End Function

Public Function Context_FillPreserve(ByVal dstContext As Long)
    cairo_fill_preserve dstContext
End Function

Public Function Context_Rectangle(ByVal dstContext As Long, ByVal dstX As Double, ByVal dstY As Double, ByVal dstWidth As Double, ByVal dstHeight As Double)
    cairo_rectangle dstContext, dstX, dstY, dstWidth, dstHeight
End Function

Public Function Context_SetAntialias(ByVal dstContext As Long, ByVal newAA As Cairo_Antialias)
    cairo_set_operator dstContext, newAA
End Function

Public Function Context_SetOperator(ByVal dstContext As Long, ByVal newOperator As Cairo_Operator)
    cairo_set_operator dstContext, newOperator
End Function

Public Function Context_SetSourcePattern(ByVal dstContext As Long, ByVal srcPattern As Long)
    cairo_set_source dstContext, srcPattern
End Function

Public Function Pattern_CreateLinearGradient(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double) As Long
    Pattern_CreateLinearGradient = cairo_pattern_create_linear(x0, y0, x1, y1)
End Function

Public Function Pattern_CreateRadialGradient(ByVal cx0 As Double, ByVal cy0 As Double, ByVal radius0 As Double, ByVal cx1 As Double, ByVal cy1 As Double, ByVal radius1 As Double) As Long
    Pattern_CreateRadialGradient = cairo_pattern_create_radial(cx0, cy0, radius0, cx1, cy1, radius1)
End Function

'Return a pattern handle to a cairo surface; this pattern can subsequently be used for painting,
' including tasks like resizing.
Public Function Pattern_GetFromSurface(ByVal srcSurface As Long) As Long
    Pattern_GetFromSurface = cairo_pattern_create_for_surface(srcSurface)
End Function

Public Sub Pattern_SetExtend(ByVal dstPattern As Long, ByVal newExtend As Cairo_Extend)
    cairo_pattern_set_extend dstPattern, newExtend
End Sub

Public Sub Pattern_SetResampleFilter(ByVal dstPattern As Long, ByVal srcFilter As Cairo_Filter)
    cairo_pattern_set_filter dstPattern, srcFilter
End Sub

Public Sub Pattern_SetStopRGB(ByVal dstPattern As Long, ByVal srcOffset As Double, ByVal srcR As Double, ByVal srcG As Double, ByVal srcB As Double)
    Const ONE_DIV_255 As Single = 1! / 255!
    cairo_pattern_add_color_stop_rgb dstPattern, srcOffset, CSng(srcR) * ONE_DIV_255, CSng(srcG) * ONE_DIV_255, CSng(srcB) * ONE_DIV_255
End Sub

Public Sub Pattern_SetStopRGBA(ByVal dstPattern As Long, ByVal srcOffset As Double, ByVal srcR As Double, ByVal srcG As Double, ByVal srcB As Double, ByVal srcA As Double)
    Const ONE_DIV_255 As Single = 1! / 255!
    cairo_pattern_add_color_stop_rgba dstPattern, srcOffset, CSng(srcR) * ONE_DIV_255, CSng(srcG) * ONE_DIV_255, CSng(srcB) * ONE_DIV_255, CSng(srcA) * ONE_DIV_255
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
        cairo_set_source_surface dstSurface.GetContextHandle(), dstSurface.GetSurfaceHandle(), 0#, 0#
        cairo_set_operator dstSurface.GetContextHandle(), co_Multiply
        
        'Paint
        PDDebug.LogAction "cairo_paint"
        cairo_paint dstSurface.GetContextHandle()
        
        'Free resources
        Set dstSurface = Nothing
        
    End If
    
End Sub

Public Sub TestOnActiveImage()

    If PDImages.IsImageActive() Then
        'Plugin_Cairo.TestPainting PDImages.GetActiveImage.GetActiveDIB
        Dim x1 As Single, y1 As Single, dstWidth As Single, dstHeight As Single
        Dim x2 As Single, y2 As Single, srcWidth As Single, srcHeight As Single
        With PDImages.GetActiveImage.GetActiveDIB
            x1 = .GetDIBWidth * 0.25
            y1 = .GetDIBHeight * 0.25
            dstWidth = .GetDIBWidth * 0.5
            dstHeight = .GetDIBHeight * 0.5
            x2 = 0
            y2 = 0
            srcWidth = .GetDIBWidth
            srcHeight = .GetDIBHeight
        End With
        
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime
        
        Plugin_Cairo.Cairo_StretchBlt PDImages.GetActiveImage.GetActiveDIB, x1, y1, dstWidth, dstHeight, PDImages.GetActiveImage.GetActiveDIB, x2, y2, srcWidth, srcHeight
        
        Debug.Print VBHacks.GetTimeDiffNowAsString(startTime)
        
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, 0
        ViewportEngine.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
        
End Sub
