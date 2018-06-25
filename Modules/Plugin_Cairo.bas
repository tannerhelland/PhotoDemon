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
    CAIRO_FORMAT_INVALID = -1
    CAIRO_FORMAT_ARGB32 = 0
    CAIRO_FORMAT_RGB24 = 1
    CAIRO_FORMAT_A8 = 2
    CAIRO_FORMAT_A1 = 3
    CAIRO_FORMAT_RGB16_565 = 4
    CAIRO_FORMAT_RGB30 = 5
End Enum

#If False Then
    Private Const CAIRO_FORMAT_INVALID = -1, CAIRO_FORMAT_ARGB32 = 0, CAIRO_FORMAT_RGB24 = 1, CAIRO_FORMAT_A8 = 2, CAIRO_FORMAT_A1 = 3, CAIRO_FORMAT_RGB16_565 = 4, CAIRO_FORMAT_RGB30 = 5
#End If

'Cairo blend operators
Private Enum Cairo_Operator
    CAIRO_OPERATOR_CLEAR
    CAIRO_OPERATOR_SOURCE
    CAIRO_OPERATOR_OVER
    CAIRO_OPERATOR_IN
    CAIRO_OPERATOR_OUT
    CAIRO_OPERATOR_ATOP
    CAIRO_OPERATOR_DEST
    CAIRO_OPERATOR_DEST_OVER
    CAIRO_OPERATOR_DEST_IN
    CAIRO_OPERATOR_DEST_OUT
    CAIRO_OPERATOR_DEST_ATOP
    CAIRO_OPERATOR_XOR
    CAIRO_OPERATOR_ADD
    CAIRO_OPERATOR_SATURATE
    CAIRO_OPERATOR_MULTIPLY
    CAIRO_OPERATOR_SCREEN
    CAIRO_OPERATOR_OVERLAY
    CAIRO_OPERATOR_DARKEN
    CAIRO_OPERATOR_LIGHTEN
    CAIRO_OPERATOR_COLOR_DODGE
    CAIRO_OPERATOR_COLOR_BURN
    CAIRO_OPERATOR_HARD_LIGHT
    CAIRO_OPERATOR_SOFT_LIGHT
    CAIRO_OPERATOR_DIFFERENCE
    CAIRO_OPERATOR_EXCLUSION
    CAIRO_OPERATOR_HSL_HUE
    CAIRO_OPERATOR_HSL_SATURATION
    CAIRO_OPERATOR_HSL_COLOR
    CAIRO_OPERATOR_HSL_LUMINOSITY
End Enum

'Exported cairo functions
Private Declare Function cairo_create Lib "cairo" (ByVal dstSurface As Long) As Long
Private Declare Sub cairo_destroy Lib "cairo" (ByVal srcContext As Long)
Private Declare Function cairo_image_surface_create_for_data Lib "cairo" (ByVal ptrToPixels As Long, ByVal pxFormat As Cairo_Format, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal imgStride As Long) As Long
Private Declare Sub cairo_paint Lib "cairo" (ByVal dstContext As Long)
Private Declare Sub cairo_set_operator Lib "cairo" (ByVal dstContext As Long, ByVal newOperator As Cairo_Operator)
Private Declare Sub cairo_set_source_rgba Lib "cairo" (ByVal dstContext As Long, ByVal srcRed As Double, ByVal srcGreen As Double, ByVal srcBlue As Double, ByVal srcAlpha As Double)
Private Declare Sub cairo_set_source_surface Lib "cairo" (ByVal dstContext As Long, ByVal srcSurface As Long, ByVal patternOriginX As Double, ByVal patternOriginY As Double)
Private Declare Sub cairo_surface_destroy Lib "cairo" (ByVal srcSurface As Long)
Private Declare Function cairo_version_string Lib "cairo" () As Long

'Persistent LoadLibrary handle; will be non-zero if cairo has been loaded.
Private m_hLibCairo As Long

'Initialize Cairo.  Do not call this until you have verified the dll's existence (typically via the PluginManager module)
Public Function InitializeCairo() As Boolean
    
    'Due to current null-pointer crashes on XP (which I have tried and failed to resolve),
    ' Cairo support is limited to Win 7+.  It's possible that the library will also run fine on Vista,
    ' but without an active test rig, I'm not going to risk it.
    If False Then
    'If OS.IsWin7OrLater Then
        
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
    VBHacks.FreeLib m_hLibCairo
End Sub

Public Function GetCairoVersion() As String
    GetCairoVersion = Strings.StringFromCharPtr(cairo_version_string(), False)
End Function

'Only works on 32-bpp DIBs at present
Public Function GetCairoSurfaceFromPDDib(ByRef srcDIB As pdDIB, ByRef dstContext As Long) As Long
    If (srcDIB.GetDIBColorDepth = 32) Then
        PDDebug.LogAction "cairo_image_surface_create_for_data"
        GetCairoSurfaceFromPDDib = cairo_image_surface_create_for_data(srcDIB.GetDIBPointer, CAIRO_FORMAT_ARGB32, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBStride)
        If (GetCairoSurfaceFromPDDib <> 0) Then
            PDDebug.LogAction "cairo_create on " & GetCairoSurfaceFromPDDib
            dstContext = cairo_create(GetCairoSurfaceFromPDDib)
            PDDebug.LogAction "dstContext = " & dstContext
        End If
    Else
        GetCairoSurfaceFromPDDib = 0
    End If
End Function

Public Sub FreeCairoContext(ByRef srcContext As Long)
    PDDebug.LogAction "cairo_destroy: " & srcContext
    If (srcContext <> 0) Then cairo_destroy srcContext
    srcContext = 0
End Sub

Public Sub FreeCairoSurface(ByRef srcSurface As Long)
    PDDebug.LogAction "cairo_surface_destroy: " & srcSurface
    If (srcSurface <> 0) Then cairo_surface_destroy srcSurface
    srcSurface = 0
End Sub

Public Sub TestPainting(ByRef srcDIB As pdDIB)
    
    Dim hSurface As Long, hContext As Long
    hSurface = Plugin_Cairo.GetCairoSurfaceFromPDDib(srcDIB, hContext)
    If (hSurface <> 0) Then
        
        'Set rendering source
        
        'Test plain color rendering:
        'PDDebug.LogAction "cairo_set_source_rgba"
        'cairo_set_source_rgba hContext, 0#, 0#, 1#, 0.5
        
        'Test full-image rendering:
        PDDebug.LogAction "cairo_set_source_surface"
        cairo_set_source_surface hContext, hSurface, 0#, 0#
        cairo_set_operator hContext, CAIRO_OPERATOR_MULTIPLY
        
        'Paint
        PDDebug.LogAction "cairo_paint"
        cairo_paint hContext
        
        'Free resources
        Plugin_Cairo.FreeCairoContext hContext
        Plugin_Cairo.FreeCairoSurface hSurface
    
    End If
    
End Sub

Public Sub TestOnActiveImage()

    If (g_OpenImageCount > 0) Then
        Plugin_Cairo.TestPainting pdImages(g_CurrentImage).GetActiveDIB
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_Layer, 0
        ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.MainCanvas(0)
    End If
        
End Sub
