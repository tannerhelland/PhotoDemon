VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdOpenRaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon OpenRaster Container and Parser
'Copyright 2019-2019 by Tanner Helland
'Created: 08/January/19
'Last updated: 09/January/19
'Last update: wrap up initial build
'
'OpenRaster is currently the "best" mechanism for sharing multi-layer images between PhotoDemon
' and other open-source photo editors (GIMP, Krita, MyPaint, etc).
'
'The OpenRaster format is described in detail at Wikipedia:
' https://en.wikipedia.org/wiki/OpenRaster
'
'The formal spec is available here (link good as of January 2019):
' https://www.openraster.org/
'
'This class requires a copy of cZipArchive, an MIT-licensed zip library by wqweto@gmail.com.
' Many thanks to wqweto for not only sharing his class under a permissive license, but also being
' very responsive to bug reports and feature requests.  An original, un-altered copy of cZipArchive
' can be downloaded from GitHub, as can its attached MIT license (link good as of January 2019):
' https://github.com/wqweto/ZipArchive
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Some of our own internal parsers can perform perf timings; enable this value to send format-specific timing reports
' to PDDebug.
Private Const GENERATE_PERF_REPORTS As Boolean = True

'cZipArchive handles all zip extraction duties
Private m_ZipArchive As cZipArchive

'When loading a (valid) ORA file, a basic header with key image attributes gets populated first.
' Note that images "in the wild" may dump a bunch of undocumented attributes into the header as well;
' I currently make no attempt to understand or preserve such settings (looking at you, MyPaint).
Private Type OpenRasterHeader
    orVersion As String 'Semantic version string
    orWidth As Long     'Image width; child layer width may be </> than this
    orHeight As Long    'Image height; child layer width may be </> than this
    orXRes As Single    'X-resolution, in PPI.  Defaults to 72.
    orYRes As Single    'X-resolution, in PPI.  Defaults to 72.
End Type

Private m_Header As OpenRasterHeader

'In a valid ORA file, layers can potentially possess many "attributes".  PhotoDemon supports most but not
' all possible properties (e.g. some blend-modes are not currently implemented).  This "PD-friendly" struct
' holds any attributes we currently support, with everything translated to its most-similar feature in PD.
Private Type OpenRasterLayer
    orlName As String
    orlX As Long
    orlY As Long
    orlOpacity As Single
    orlVisibility As Boolean
    orlBlendMode As PD_BlendMode
    orlSelected As Boolean      'Not in the spec, but appears to describe a currently active layer in Krita;
                                ' we use it to set the active layer, if found.
    orlSourceFile As String     'Used to know which PNG file to extract from the resources section,
                                ' but we don't store this attribute outside of this class
End Type

'When validating an ORA file, we store the original filename that was validated; if the validation succeeds,
' we can reuse this same m_ZipArchive instance for actual parsing duties.
Private m_Filename As String

'Validate a source filename as ORA format.  Validation *does* touch the file - we actually open the first
' file in the archive to make sure it meets the required mimetype descriptor.
Friend Function IsFileORA(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True) As Boolean

    Dim potentiallyORA As Boolean
    potentiallyORA = True
    
    'Perf reports are optional, and they rely on some PD-specific timing components.  Feel free to
    ' comment these lines out as necessary.
    Dim startTime As Currency
    If GENERATE_PERF_REPORTS Then VBHacks.GetHighResTime startTime
    
    'Check extension up front, if requested
    If requireValidFileExtension Then potentiallyORA = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "ora", True)
    
    'Proceed with deeper validation as necessary
    If potentiallyORA Then
        
        'Attempt to load the archive
        potentiallyORA = m_ZipArchive.OpenArchive(srcFilename)
        If potentiallyORA Then
        
            'The file appears to be a valid zip archive.  Retrieve a list of included filenames;
            ' we're looking for "mimetype" as the first entry.
            Const IDX_FILENAME As Long = 0
            potentiallyORA = Strings.StringsEqual(CStr(m_ZipArchive.FileInfo(0)(IDX_FILENAME)), "mimetype", False)
            
            'If the mimetype file is present, extract it and ensure the contents are, per the spec:
            ' "...the string "image/openraster", with no whitespace or trailing newline."
            If potentiallyORA Then
                
                Dim utf8Bytes() As Byte
                potentiallyORA = m_ZipArchive.Extract(utf8Bytes, "mimetype")
                If potentiallyORA Then potentiallyORA = Strings.StringsEqual(Strings.StringFromUTF8(utf8Bytes), "image/openraster")
                
                'On success, don't free the cZipArchive instance; we want to reuse the current instance
                ' for actual file import.  Instead, note the successfully validated filename so that the
                ' loader knows to skip validation steps.
                If potentiallyORA Then m_Filename = srcFilename
                
            End If
        
        End If
        
    End If
    
    If GENERATE_PERF_REPORTS Then InternalWarning "IsFileORA", "file validation timing report: " & VBHacks.GetTimeDiffNowAsString(startTime)
    
    IsFileORA = potentiallyORA
    
End Function

'Before loading an actual ORA file, consider running IsFileORA(), above - this will validate the file for you,
' and you can avoid calling this function for non-ORA files.
Friend Function LoadORA(ByRef srcFilename As String, ByRef dstImage As pdImage) As Boolean
    
    LoadORA = False
    
    On Error GoTo CouldNotLoadFile
    
    'If we haven't validated the target file, do so now
    If Strings.StringsNotEqual(m_Filename, srcFilename) Then
        If (Not Me.IsFileORA(m_Filename)) Then
            InternalError "LoadORA", "target file isn't in OpenRaster format"
            Exit Function
        End If
    End If
    
    Dim startTime As Currency, firstTime As Currency
    Dim zipTime As Double, pngTime As Double, postTime As Double
    If GENERATE_PERF_REPORTS Then
        VBHacks.GetHighResTime firstTime
        VBHacks.GetHighResTime startTime
    End If
    
    'The validator will have already loaded the target file, so we can proceed with extraction.
    
    'The first thing we want to do is extract the stack.xml file.  This contains crucial attributes like
    ' the image's width and height, which we need to know before we start parsing individual layers.
    Dim utf8Bytes() As Byte
    If m_ZipArchive.Extract(utf8Bytes, "stack.xml") Then
        
        'Validate and load the XML
        Dim xmlEngine As MSXML2.DOMDocument
        Set xmlEngine = New MSXML2.DOMDocument
        xmlEngine.async = False
        xmlEngine.validateOnParse = True
        
        Dim srcXML As String
        srcXML = Strings.StringFromUTF8(utf8Bytes)
        
        If xmlEngine.loadXML(srcXML) Then
        
            'The root node should be <image>.
            Dim xmlRoot As MSXML2.IXMLDOMNode
            Set xmlRoot = xmlEngine.documentElement
            If (xmlRoot.Attributes.Length > 0) Then
                
                'Prep a default header; we will overwrite these attributes if/when we encounter them
                With m_Header
                    .orVersion = vbNullString
                    .orWidth = 0
                    .orHeight = 0
                    .orXRes = 72    'x/y res default to 72 per the spec; these attributes are optional, however
                    .orYRes = 72
                End With
                
                'The final validation we need to perform involves checking the root image's attributes.
                ' We need to make sure basic things are intact, like non-zero width/height
                Dim i As Long, curNodeName As String, curNodeValue As String
                For i = 0 To xmlRoot.Attributes.Length - 1
                    
                    curNodeName = LCase$(xmlRoot.Attributes(i).nodeName)
                    curNodeValue = xmlRoot.Attributes(i).nodeValue
                    
                    Select Case curNodeName
                        Case "version"
                            m_Header.orVersion = curNodeValue
                        Case "w"
                            m_Header.orWidth = CLng(curNodeValue)
                        Case "h"
                            m_Header.orHeight = CLng(curNodeValue)
                        Case "xres"
                            m_Header.orXRes = TextSupport.CDblCustom(curNodeValue)
                        Case "yres"
                            m_Header.orYRes = TextSupport.CDblCustom(curNodeValue)
                        'The spec doesn't define any other image attributes, but who knows what we'll
                        ' encounter in the wild!
                        Case Else
                            InternalError "LoadORA", "Unknown <image> attribute encountered: " & curNodeName & "(" & curNodeValue & ")"
                    End Select
                    
                Next i
                
                'Make sure we found a valid width and height.  (The spec wants a version check as well,
                ' but this seems overkill if everything else is valid.)
                If (m_Header.orWidth > 0) And (m_Header.orHeight > 0) Then
                
                    'The image passes validation!  This appears to be a valid OpenRaster file.
                    With m_Header
                        dstImage.Width = .orWidth
                        dstImage.Height = .orHeight
                        dstImage.SetDPI .orXRes, .orYRes
                    End With
                    
                    ' We can now begin the messy business of traversing the rest of the stack.xml file
                    ' and extracting associated PNG data as we go.
                    
                    'TODO: the spec allows for layer grouping ("stacks"), which PhotoDemon does not currently support.
                    ' There's no easy workaround for this until we implement groups, and I have no ETA for this at present.
                    ' As such, I currently ignore stacks and simply iterate layer entries.  This will still generate
                    ' a correct image for the majority of "in the wild" ORA files.
                    
                    'Because of this, instead of traversing the DOM we can simply grab a list of layer nodes.
                    ' Without grouping support, the only thing we really care about is order, not nesting.
                    Dim listOfLayers As MSXML2.IXMLDOMNodeList
                    Set listOfLayers = xmlEngine.getElementsByTagName("layer")
                    
                    'Traverse the list of layers in *reverse* order, as the last layer in the file is the base layer,
                    ' while the first layer is the top-most one.
                    Dim layerProps() As OpenRasterLayer
                    ReDim layerProps(0 To listOfLayers.Length - 1) As OpenRasterLayer
                    
                    Dim j As Long
                    For i = listOfLayers.Length - 1 To 0 Step -1
                    
                        With layerProps(i)
                        
                            For j = 0 To listOfLayers(i).Attributes.Length - 1
                                
                                curNodeName = LCase$(listOfLayers(i).Attributes(j).nodeName)
                                curNodeValue = listOfLayers(i).Attributes(j).nodeValue
                                
                                Select Case curNodeName
                                    
                                    'Spec-defined attributes follow
                                    Case "name"
                                        .orlName = curNodeValue
                                    Case "opacity"
                                        .orlOpacity = TextSupport.CDblCustom(curNodeValue) * 100#
                                    Case "src"
                                        .orlSourceFile = curNodeValue
                                    Case "visibility"
                                        .orlVisibility = Strings.StringsEqual(curNodeValue, "visible", True)
                                    Case "x"
                                        .orlX = CLng(TextSupport.CDblCustom(curNodeValue))
                                    Case "y"
                                        .orlY = CLng(TextSupport.CDblCustom(curNodeValue))
                                    Case "composite-op"
                                        .orlBlendMode = GetBlendModeFromSVGOp(curNodeValue)
                                    
                                    'Non-spec defined attributes found in the wild follow
                                    
                                    'Krita uses "selected" to mark active layers; PD only supports one "selected"
                                    ' layer at a time, so we'll end up grabbing the *lowest in the stack* selected
                                    ' layer for the time being.
                                    Case "selected"
                                        .orlSelected = Strings.StringsEqual(curNodeValue, "true", True)
                                    
                                    'The spec doesn't define any other layer attributes, but who knows what we'll
                                    ' encounter in the wild!
                                    Case Else
                                        InternalError "LoadORA", "Unknown <layer> attribute encountered: " & curNodeName & "(" & curNodeValue & ")"
                                End Select
                                
                            Next j
                        End With
                        
                    Next i
                    
                    If GENERATE_PERF_REPORTS Then
                        zipTime = 0#
                        pngTime = 0#
                        postTime = 0#
                    End If
                    
                    'We now have all the information we need to reconstruct individual image layers.  Note that the OpenRaster spec
                    ' only allows individual layers to be embedded PNG images, but some software (MyPaint) may also embed SVG layers.
                    ' We can't parse SVGs at present, so we're SOL on such layers.
                    Dim cPNG As pdPNG, tmpPngBytes() As Byte
                    Dim tmpDIB As pdDIB, tmpLayer As pdLayer, newLayerID As Long, selLayerIndex As Long
                    For i = listOfLayers.Length - 1 To 0 Step -1
                        
                        If GENERATE_PERF_REPORTS Then VBHacks.GetHighResTime startTime
                        
                        'Before creating a layer, we want to first attempt to retrieve the layer's PNG data.
                        ' (If this fails, we'll discard the layer entirely.)
                        If m_ZipArchive.Extract(tmpPngBytes, layerProps(i).orlSourceFile) Then
                            
                            If GENERATE_PERF_REPORTS Then
                                zipTime = zipTime + VBHacks.GetTimerDifferenceNow(startTime)
                                VBHacks.GetHighResTime startTime
                            End If
                            
                            Set tmpDIB = New pdDIB
                            Set cPNG = New pdPNG
                            If (cPNG.LoadPNG_Simple(vbNullString, Nothing, tmpDIB, False, VarPtr(tmpPngBytes(0)), UBound(tmpPngBytes) + 1) = png_Success) Then
                                
                                If GENERATE_PERF_REPORTS Then
                                    pngTime = pngTime + VBHacks.GetTimerDifferenceNow(startTime)
                                    VBHacks.GetHighResTime startTime
                                End If
                                
                                'Because color-management has already been handled (if applicable), this is a great time to premultiply alpha
                                tmpDIB.SetAlphaPremultiplication True
                                
                                'Prep a new layer object and initialize it with the image bits we've retrieved
                                newLayerID = dstImage.CreateBlankLayer()
                                Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                                tmpLayer.InitializeNewLayer PDL_IMAGE, layerProps(i).orlName, tmpDIB
                                
                                'Fill in any remaining layer properties
                                With layerProps(i)
                                    tmpLayer.SetLayerBlendMode .orlBlendMode
                                    tmpLayer.SetLayerOpacity .orlOpacity
                                    tmpLayer.SetLayerOffsetX .orlX
                                    tmpLayer.SetLayerOffsetY .orlY
                                    tmpLayer.SetLayerVisibility .orlVisibility
                                    If .orlSelected Then selLayerIndex = i
                                End With
                                
                                'Notify the layer of new changes, so it knows to regenerate internal caches on next access
                                tmpLayer.NotifyOfDestructiveChanges
                                
                                If GENERATE_PERF_REPORTS Then postTime = postTime + VBHacks.GetTimerDifferenceNow(startTime)
                                
                            Else
                                InternalError "LoadORA", "Retrieved bytes for layer #" & CStr(i) & " were not PNG format; filename is: " & layerProps(i).orlSourceFile
                            End If
                            
                        Else
                            InternalError "LoadORA", "Could not extract PNG bytes for layer #" & CStr(i) & ": " & layerProps(i).orlSourceFile
                        End If
                        
                    Next i
                    
                    'All layers have been iterated.  If the target image contains at least one valid layer,
                    ' consider this a successful load.
                    LoadORA = (dstImage.GetNumOfLayers > 0)
                
                Else
                    InternalError "LoadORA", "Image width and/or height invalid (" & m_Header.orWidth & "x" & m_Header.orHeight & ")"
                End If
                
            Else
                InternalError "LoadORA", "root image node provides no attributes"
                LoadORA = False
            End If
            
        Else
            InternalError "LoadORA", "stack.xml didn't validate"
            LoadORA = False
        End If
    
    Else
        'On failure, there's nothing we can do - we lack enough information to construct even a placeholder image
        InternalError "LoadORA", "no stack.xml in file"
        LoadORA = False
    End If
    
    If GENERATE_PERF_REPORTS Then
        InternalWarning "LoadORA", "Total time to load OpenRaster file: " & VBHacks.GetTimeDiffNowAsString(firstTime)
        InternalWarning "LoadORA", "Time spent in zip extraction: " & Format$(zipTime * 1000#, "0.0") & " ms"
        InternalWarning "LoadORA", "Time spent in PNG parsing: " & Format$(pngTime * 1000#, "0.0") & " ms"
        InternalWarning "LoadORA", "Time spent in post-processing: " & Format$(postTime * 1000#, "0.0") & " ms"
    End If
    
    Exit Function
    
CouldNotLoadFile:
    InternalError "LoadORA", "Internal VB error #" & Err.Number & ": " & Err.Description
    LoadORA = False

End Function

Private Function GetBlendModeFromSVGOp(ByRef srcString As String) As PD_BlendMode

    Select Case LCase$(srcString)
        
        'Normal / Source Over
        Case "svg:src-over"
            GetBlendModeFromSVGOp = BL_NORMAL
            
        'Multiply / Source Over
        Case "svg:multiply"
            GetBlendModeFromSVGOp = BL_MULTIPLY
        
        'Screen / Source Over
        Case "svg:screen"
            GetBlendModeFromSVGOp = BL_SCREEN
        
        'Overlay / Source Over
        Case "svg:overlay"
            GetBlendModeFromSVGOp = BL_OVERLAY
        
        'Darken / Source Over
        Case "svg:darken"
            GetBlendModeFromSVGOp = BL_DARKEN
        
        'Lighten / Source Over
        Case "svg:lighten"
            GetBlendModeFromSVGOp = BL_LIGHTEN
            
        'Color Dodge / Source Over
        Case "svg:color-dodge"
            GetBlendModeFromSVGOp = BL_COLORDODGE
        
        'Color Burn / Source Over
        Case "svg:color-burn"
            GetBlendModeFromSVGOp = BL_COLORBURN
        
        'Hard Light / Source Over
        Case "svg:hard-light"
            GetBlendModeFromSVGOp = BL_HARDLIGHT
        
        'Soft Light / Source Over
        Case "svg:soft-light"
            GetBlendModeFromSVGOp = BL_SOFTLIGHT
        
        'Difference / Source Over
        Case "svg:difference"
            GetBlendModeFromSVGOp = BL_DIFFERENCE
        
        'Color / Source Over
        Case "svg:color"
            GetBlendModeFromSVGOp = BL_COLOR
        
        'Luminosity / Source Over
        Case "svg:luminosity"
            GetBlendModeFromSVGOp = BL_LUMINOSITY
        
        'Hue / Source Over
        Case "svg:hue"
            GetBlendModeFromSVGOp = BL_HUE
        
        'Saturation / Source Over
        Case "svg:saturation"
            GetBlendModeFromSVGOp = BL_SATURATION
        
        'Normal / Lighter
        Case "svg:plus"
            InternalError "GetBlendModeFromSVGOp", "unsupported blend mode: " & srcString
            GetBlendModeFromSVGOp = BL_NORMAL
        
        'Normal / Destination In
        Case "svg:dst-in"
            InternalError "GetBlendModeFromSVGOp", "unsupported blend mode: " & srcString
            GetBlendModeFromSVGOp = BL_NORMAL
        
        'Normal / Destination Out
        Case "svg:dst-out"
            InternalError "GetBlendModeFromSVGOp", "unsupported blend mode: " & srcString
            GetBlendModeFromSVGOp = BL_NORMAL
        
        'Normal / Source Atop
        Case "svg:src-atop"
            InternalError "GetBlendModeFromSVGOp", "unsupported blend mode: " & srcString
            GetBlendModeFromSVGOp = BL_NORMAL
        
        'Normal / Destination Atop
        Case "svg:dst-atop"
            InternalError "GetBlendModeFromSVGOp", "unsupported blend mode: " & srcString
            GetBlendModeFromSVGOp = BL_NORMAL
        
        Case Else
            InternalError "GetBlendModeFromSVGOp", "unknown blend mode: " & srcString
            GetBlendModeFromSVGOp = BL_NORMAL
    
    End Select

End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdOpenRaster." & funcName & "() reported an error on file """ & m_Filename & """: " & errDescription
    Else
        Debug.Print "pdOpenRaster." & funcName & "() reported an error on file """ & m_Filename & """: " & errDescription
    End If
End Sub

Private Sub InternalWarning(ByRef funcName As String, ByRef warnText As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdOpenRaster." & funcName & "() warned: " & warnText
    Else
        Debug.Print "pdOpenRaster." & funcName & "() warned: " & warnText
    End If
End Sub

Private Sub Class_Initialize()
    Set m_ZipArchive = New cZipArchive
End Sub
