Attribute VB_Name = "Plugin_resvg"
'***************************************************************************
'resvg Library Interface (SVG import)
'Copyright 2022-2022 by Tanner Helland
'Created: 28/February/22
'Last updated: 01/March/22
'Last update: wrap up initial build
'
'Per its documentation (available at https://github.com/RazrFalcon/resvg), resvg is...
'
' "...an SVG rendering library.
' It can be used as a Rust library, as a C library and as a CLI application to render static SVG files.
' The core idea is to make a fast, small, portable SVG library with an aim to support the whole SVG spec."
'
'Yevhenii Reizner is the author of resvg.  resvg is MPL-licensed and actively maintained.
' The copy of resvg.dll that ships with PhotoDemon is based on the 0.22.0 release and built against the
' i686-pc-windows-msvc rust target (for XP support).  It *must* be hand-edited to export stdcall funcs.
' (Normally I just use cdecl via DispCallFunc, but resvg returns some custom types that don't work
' with DispCallFunc - so manually building against stdcall is necessary.)  Note that some function decs
' must also be rewritten to pass UDTs as references instead of values, as required by VB6.
'
'A full copy of the resvg license is available here:
' https://github.com/RazrFalcon/resvg/blob/master/LICENSE.txt
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Information on individual resvg calls can be saved to the debug log via this constant;
' please DISABLE in production builds
Private Const SVG_DEBUG_VERBOSE As Boolean = False

Private Enum resvg_result
    'Everything is ok.
    RESVG_OK = 0
    'Only UTF-8 content are supported.
    RESVG_ERROR_NOT_AN_UTF8_STR = 1
    'Failed to open the provided file.
    RESVG_ERROR_FILE_OPEN_FAILED = 2
    'Compressed SVG must use the GZip algorithm.
    RESVG_ERROR_MALFORMED_GZIP = 3
    'We do not allow SVG with more than 1_000_000 elements for security reasons.
    RESVG_ERROR_ELEMENTS_LIMIT_REACHED = 4
    'SVG doesn't have a valid size.
    '   (Occurs when width and/or height are <= 0.)
    '   (Also occurs if width, height and viewBox are not set.)
    RESVG_ERROR_INVALID_SIZE = 5
    'Failed to parse an SVG data.
    RESVG_ERROR_PARSING_FAILED = 6
End Enum

#If False Then
    Private Const RESVG_OK = 0, RESVG_ERROR_NOT_AN_UTF8_STR = 1, RESVG_ERROR_FILE_OPEN_FAILED = 2, RESVG_ERROR_MALFORMED_GZIP = 3, RESVG_ERROR_ELEMENTS_LIMIT_REACHED = 4, RESVG_ERROR_INVALID_SIZE = 5, RESVG_ERROR_PARSING_FAILED = 6
#End If

'A "fit to" type.
' (All types produce proportional scaling.)
Private Enum resvg_fit_to_type
    'Use an original image size.
    RESVG_FIT_TO_TYPE_ORIGINAL
    'Fit an image to a specified width.
    RESVG_FIT_TO_TYPE_WIDTH
    'Fit an image to a specified height.
    RESVG_FIT_TO_TYPE_HEIGHT
    'Zoom an image using scaling factor.
    RESVG_FIT_TO_TYPE_ZOOM
End Enum

#If False Then
    Private Const RESVG_FIT_TO_TYPE_ORIGINAL = 0, RESVG_FIT_TO_TYPE_WIDTH = 0, RESVG_FIT_TO_TYPE_HEIGHT = 0, RESVG_FIT_TO_TYPE_ZOOM = 0
#End If

'An image rendering method.
Private Enum resvg_image_rendering
    RESVG_IMAGE_RENDERING_OPTIMIZE_QUALITY
    RESVG_IMAGE_RENDERING_OPTIMIZE_SPEED
End Enum

#If False Then
    Private Const RESVG_IMAGE_RENDERING_OPTIMIZE_QUALITY = 0, RESVG_IMAGE_RENDERING_OPTIMIZE_SPEED = 0
#End If

'A shape rendering method.
Private Enum resvg_shape_rendering
    RESVG_SHAPE_RENDERING_OPTIMIZE_SPEED
    RESVG_SHAPE_RENDERING_CRISP_EDGES
    RESVG_SHAPE_RENDERING_GEOMETRIC_PRECISION
End Enum

#If False Then
    Private Const RESVG_SHAPE_RENDERING_OPTIMIZE_SPEED = 0, RESVG_SHAPE_RENDERING_CRISP_EDGES = 0, RESVG_SHAPE_RENDERING_GEOMETRIC_PRECISION = 0
#End If

'A text rendering method.
Private Enum resvg_text_rendering
    RESVG_TEXT_RENDERING_OPTIMIZE_SPEED
    RESVG_TEXT_RENDERING_OPTIMIZE_LEGIBILITY
    RESVG_TEXT_RENDERING_GEOMETRIC_PRECISION
End Enum

#If False Then
    Private Const RESVG_TEXT_RENDERING_OPTIMIZE_SPEED = 0, RESVG_TEXT_RENDERING_OPTIMIZE_LEGIBILITY = 0, RESVG_TEXT_RENDERING_GEOMETRIC_PRECISION = 0
#End If

'A 2D transform representation.
Private Type resvg_transform
    a As Double
    b As Double
    c As Double
    d As Double
    e As Double
    f As Double
End Type

'A size representation.
' (Width and height are guaranteed to be > 0.)
Private Type resvg_size
    svg_width As Double
    svg_height As Double
End Type

'A rectangle representation.
' (Width *and* height are guarantee to be > 0.)
Private Type resvg_rect
    x As Double
    y As Double
    Width As Double
    Height As Double
End Type

'A path bbox representation.
' (Width *or* height are guarantee to be > 0.)
Private Type resvg_path_bbox
    x As Double
    y As Double
    Width As Double
    Height As Double
End Type

'A "fit to" property.
Private Type resvg_fit_to
    'A fit type.
    fit_type As resvg_fit_to_type
    'Fit to value
    '* Not used by RESVG_FIT_TO_ORIGINAL.
    '* Must be >= 1 for RESVG_FIT_TO_WIDTH and RESVG_FIT_TO_HEIGHT.
    '* Must be > 0 for RESVG_FIT_TO_ZOOM.
    fit_value As Single
End Type

Private Declare Function resvg_transform_identity Lib "resvg" () As resvg_transform
Private Declare Sub resvg_init_log Lib "resvg" ()
Private Declare Function resvg_options_create Lib "resvg" () As Long
Private Declare Sub resvg_options_set_resources_dir Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_dpi Lib "resvg" (ByVal resvg_options As Long, ByVal newDPI As Double)
Private Declare Sub resvg_options_set_font_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_font_size Lib "resvg" (ByVal resvg_options As Long, ByVal newSize As Double)
Private Declare Sub resvg_options_set_serif_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_sans_serif_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_cursive_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_fantasy_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_monospace_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_languages Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Languages As Long)
Private Declare Sub resvg_options_set_shape_rendering_mode Lib "resvg" (ByVal resvg_options As Long, ByVal newMode As resvg_shape_rendering)
Private Declare Sub resvg_options_set_text_rendering_mode Lib "resvg" (ByVal resvg_options As Long, ByVal newMode As resvg_text_rendering)
Private Declare Sub resvg_options_set_image_rendering_mode Lib "resvg" (ByVal resvg_options As Long, ByVal newMode As resvg_image_rendering)
Private Declare Sub resvg_options_set_keep_named_groups Lib "resvg" (ByVal resvg_options As Long, ByVal keepBool As Long)
Private Declare Sub resvg_options_load_font_data Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToData As Long, ByVal sizeOfData As Long)
Private Declare Function resvg_options_load_font_file Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8FilePath As Long) As resvg_result
Private Declare Sub resvg_options_load_system_fonts Lib "resvg" (ByVal resvg_options As Long)
Private Declare Sub resvg_options_destroy Lib "resvg" (ByVal resvg_options As Long)
Private Declare Function resvg_parse_tree_from_file Lib "resvg" (ByVal ptrToConstUtf8FilePath As Long, ByVal resvg_options As Long, ByRef resvg_render_tree As Long) As Long
Private Declare Function resvg_parse_tree_from_data Lib "resvg" (ByVal ptrToData As Long, ByVal sizeOfData As Long, ByVal resvg_options As Long, ByRef resvg_render_tree As Long) As Long
Private Declare Function resvg_is_image_empty Lib "resvg" (ByVal resvg_render_tree As Long) As Long
Private Declare Function resvg_get_image_size Lib "resvg" (ByVal resvg_render_tree As Long) As resvg_size
Private Declare Function resvg_get_image_viewbox Lib "resvg" (ByVal resvg_render_tree As Long) As resvg_rect
Private Declare Function resvg_get_imgae_bbox Lib "resvg" (ByVal resvg_render_tree As Long, ByRef dst_resvg_rect As resvg_rect) As Long
Private Declare Function resvg_node_exists Lib "resvg" (ByVal resvg_render_tree As Long, ByVal ptrToConstUtf8ID As Long) As Long
Private Declare Function resvg_get_node_transform Lib "resvg" (ByVal resvg_render_tree As Long, ByVal ptrToConstUtf8ID As Long, ByRef dst_resvg_transform As resvg_transform) As Long
Private Declare Function resvg_get_node_bbox Lib "resvg" (ByVal resvg_render_tree As Long, ByVal ptrToConstUtf8ID As Long, ByRef dst_resvg_path_bbox As resvg_path_bbox) As Long
Private Declare Sub resvg_tree_destroy Lib "resvg" (ByVal resvg_render_tree As Long)
Private Declare Sub resvg_render Lib "resvg" (ByVal resvg_render_tree As Long, ByRef fit_to As resvg_fit_to, ByRef srcTransform As resvg_transform, ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, ByVal ptrToSurface As Long)
Private Declare Sub resvg_render_node Lib "resvg" (ByVal resvg_render_tree As Long, ByVal ptrToConstUtf8ID As Long, ByVal fit_to As resvg_fit_to, ByVal srcTransform As resvg_transform, ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, ByVal ptrToSurface As Long)

'Library handle will be non-zero if all required dll(s) are available;
' you can also forcibly override the "availability" state by setting m_LibAvailable to FALSE.
' (This effectively disables run-time support in the UI.)
Private m_LibHandle As Long, m_LibAvailable As Boolean

'Forcibly disable this library at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetVersion() As String
    
    'resvg does not provide an externally accessible version string by default.
    ' I do not expect users to custom-build it, so we return a hard-coded version
    ' against the copy supplied with a default PD install.
    GetVersion = "0.22.0"
    
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim strLibPath As String
    strLibPath = pathToDLLFolder & "resvg.dll"
    m_LibHandle = VBHacks.LoadLib(strLibPath)
    m_LibAvailable = (m_LibHandle <> 0)
    InitializeEngine = m_LibAvailable
    
    If InitializeEngine Then
        PDDebug.LogAction "SVG support enabled"
    Else
        PDDebug.LogAction "WARNING!  LoadLibraryW failed to load resvg.  Last DLL error: " & Err.LastDllError
    End If
    
End Function

'Simple extension check on SVG data.  resvg will return pass/fail on parsing, but we currently also
' limit file extension as a perf-friendly way to avoid touching resvg if we don't need to.
Public Function IsFileSVGCandidate(ByRef imagePath As String) As Boolean
    If Plugin_resvg.IsResvgEnabled Then
        IsFileSVGCandidate = Strings.StringsEqual(Right$(imagePath, 3), "svg", True) Or Strings.StringsEqual(Right$(imagePath, 4), "svgz", True)
    Else
        IsFileSVGCandidate = False
    End If
End Function

Public Function IsResvgEnabled() As Boolean
    IsResvgEnabled = m_LibAvailable
End Function

'Given a source SVG file, attempt to load it into a target pdImage/pdDIB
Public Function LoadSVG_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    LoadSVG_FromFile = False
    
    If SVG_DEBUG_VERBOSE Then PDDebug.LogAction "Handing svg parsing duties over to resvg..."
    
    'In the future, we could look at initializing logging - but data is passed to the default stderr instance,
    ' which isn't ideal...
    'resvg_init_log();
    
    'Create a blank resvg options object
    Dim svgOptions As Long
    svgOptions = resvg_options_create()
    
    'Potentially expose this via UI, but we also need to initialize a system font list
    ' (in case the SVG file embeds font data).  Note that there is a (potentially significant)
    ' perf penalty to this.
    resvg_options_load_system_fonts svgOptions
    If SVG_DEBUG_VERBOSE Then PDDebug.LogAction "System fonts ready for any embedded SVG text data"
    
    'Pre-set any other options here, as desired...
    
    'Create a blank resvg render tree pointer, and note that this is the first call where we get
    ' an actual success/fail return
    Dim svgResult As resvg_result, svgTree As Long
    
    Dim utf8path() As Byte, utf8Len As Long
    Strings.UTF8FromString srcFile, utf8path, utf8Len
    svgResult = resvg_parse_tree_from_file(VarPtr(utf8path(0)), svgOptions, svgTree)
    
    If (svgResult = RESVG_OK) Then
        If SVG_DEBUG_VERBOSE Then PDDebug.LogAction "Successfully retrieved SVG tree: " & svgTree
    Else
        InternalError vbNullString, svgResult
        LoadSVG_FromFile = False
        GoTo SafeCleanup
    End If
    
    'Can safely delete options here, I think?  (sample cairo file demonstrates this)
    
    'Retrieve image size and convert to integers (to be used as surface dimensions)
    Dim imgSize As resvg_size
    imgSize = resvg_get_image_size(svgTree)
    
    Dim intWidth As Long, intHeight As Long
    intWidth = Int(imgSize.svg_width)
    intHeight = Int(imgSize.svg_height)
    
    'Prep the target DIB
    Set dstDIB = New pdDIB
    If dstDIB.CreateBlank(intWidth, intHeight, 32, 0, 0) Then
        
        'SVG render will always be premultiplied
        dstDIB.SetInitialAlphaPremultiplicationState True
        
        'Specify fitting behavior (should make this user-controlled in the future)
        Dim fitBehavior As resvg_fit_to
        fitBehavior.fit_type = RESVG_FIT_TO_TYPE_ORIGINAL
        fitBehavior.fit_value = 1!
        
        'Render!
        Dim idMatrix As resvg_transform
        idMatrix = resvg_transform_identity()
        resvg_render svgTree, fitBehavior, idMatrix, intWidth, intHeight, dstDIB.GetDIBPointer()
        PDDebug.LogAction "Finished render"
        
        'Finally, we need to swizzle RGBA order to BGRA order
        DIBs.SwizzleBR dstDIB
        
        LoadSVG_FromFile = True
        
    Else
        InternalError "couldn't initialize DIB (probably OOM)"
    End If
    
SafeCleanup:

    'On success OR failure, free any opaque references
    If (svgOptions <> 0) Then resvg_options_destroy svgOptions
    If (svgTree <> 0) Then resvg_tree_destroy svgTree

End Function

Public Sub ReleaseEngine()
    If (m_LibHandle <> 0) Then
        VBHacks.FreeLib m_LibHandle
        m_LibHandle = 0
    End If
End Sub

Private Sub InternalError(ByVal errString As String, Optional ByVal faultyReturnCode As resvg_result = RESVG_OK)
    If (faultyReturnCode <> RESVG_OK) Then
        
        Select Case faultyReturnCode
            Case RESVG_ERROR_NOT_AN_UTF8_STR
                errString = "Only UTF-8 content supported"
            Case RESVG_ERROR_FILE_OPEN_FAILED
                errString = "Failed to open file"
            Case RESVG_ERROR_MALFORMED_GZIP
                errString = "Compressed SVG problem (must use gzip)"
            Case RESVG_ERROR_ELEMENTS_LIMIT_REACHED
                errString = "SVG elements limited to 1,000,000"
            Case RESVG_ERROR_INVALID_SIZE
                errString = "invalid size (width/height < 0)"
            Case RESVG_ERROR_PARSING_FAILED
                errString = "SVG parse failed"
        End Select
        
        PDDebug.LogAction "WARNING! resvg error (" & faultyReturnCode & "): " & errString, PDM_External_Lib
    Else
        PDDebug.LogAction "WARNING! resvg problem: " & errString, PDM_External_Lib
    End If
End Sub
